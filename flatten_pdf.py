import subprocess, sys, os, shutil, argparse, tempfile

GS_FALLBACK = r"C:\Program Files\gs\gs10.06.0\bin\gswin64c.exe"

def gs_exe():
    found = shutil.which("gswin64c")
    if found:
        return found
    if os.path.exists(GS_FALLBACK):
        return GS_FALLBACK
    raise FileNotFoundError("gswin64c not found on PATH or at " + GS_FALLBACK)

def _blank_form(writer):
    from pypdf.generic import DecodedStreamObject, NameObject, ArrayObject, FloatObject
    form = DecodedStreamObject()
    form[NameObject('/Type')]    = NameObject('/XObject')
    form[NameObject('/Subtype')] = NameObject('/Form')
    form[NameObject('/BBox')]    = ArrayObject([FloatObject(0), FloatObject(0),
                                                FloatObject(1), FloatObject(1)])
    form._data = b''
    return writer._add_object(form)

def _annotation_rect(page):
    from pypdf.generic import NameObject
    annots = page.get('/Annots')
    if not annots:
        return None
    for a in annots:
        obj = a.get_object() if hasattr(a, 'get_object') else a
        if obj.get('/Subtype') in (NameObject('/Square'), NameObject('/Rectangle')):
            r = obj['/Rect']
            return tuple(float(v) for v in r)
    return None

def _prepend_clip(page, writer, crop):
    from pypdf.generic import DecodedStreamObject, ArrayObject, NameObject
    x0, y0, x1, y1 = crop
    clip_bytes = f"{x0} {y0} {x1-x0} {y1-y0} re W n\n".encode()
    clip_obj = DecodedStreamObject()
    clip_obj._data = clip_bytes
    clip_ref = writer._add_object(clip_obj)
    contents_key = NameObject('/Contents')
    existing = page.get('/Contents')
    if existing is None:
        page[contents_key] = clip_ref
    elif isinstance(existing, ArrayObject):
        existing.insert(0, clip_ref)
    else:
        page[contents_key] = ArrayObject([clip_ref, existing])

def sanitize(reader, dst, no_images=False, crop=None):
    from pypdf import PdfWriter
    from pypdf.generic import NameObject, RectangleObject

    writer = PdfWriter()
    writer.append(reader)
    page = writer.pages[0]

    if no_images:
        resources = page['/Resources']
        if hasattr(resources, 'get_object'):
            resources = resources.get_object()
        xobj = resources.get('/XObject')
        if xobj is not None:
            if hasattr(xobj, 'get_object'):
                xobj = xobj.get_object()
            replaced = 0
            for k in list(xobj.keys()):
                ref = xobj[k]
                obj = ref.get_object() if hasattr(ref, 'get_object') else ref
                if hasattr(obj, 'get') and obj.get('/Subtype') == NameObject('/Image'):
                    xobj[k] = _blank_form(writer)
                    replaced += 1
            print(f"  Replaced {replaced} image(s) with blank forms")

    if crop is not None:
        page.cropbox = RectangleObject(crop)
        print(f"  CropBox: {crop}")

    with open(dst, 'wb') as f:
        writer.write(f)

def filter_tiny_paths(src, dst, min_size):
    try:
        import pikepdf
    except ImportError:
        raise ImportError("pikepdf required for --min-path-size. Install: pip install pikepdf")

    PATH_BUILD = frozenset({'m', 'l', 'c', 'v', 'y', 're', 'h', 'W', 'W*'})
    PATH_PAINT = frozenset({'S', 's', 'f', 'F', 'f*', 'B', 'B*', 'b', 'b*', 'n'})
    dropped = 0

    with pikepdf.open(src) as pdf:
        for page in pdf.pages:
            instructions = list(pikepdf.parse_content_stream(page))
            out, pending, xs, ys = [], [], [], []

            for operands, operator in instructions:
                op = str(operator)

                if op in PATH_BUILD:
                    pending.append((operands, operator))
                    ops = [float(o) for o in operands]
                    if op in ('m', 'l'):
                        xs.append(ops[0]); ys.append(ops[1])
                    elif op == 're':
                        xs += [ops[0], ops[0] + ops[2]]; ys += [ops[1], ops[1] + ops[3]]
                    elif op == 'c':
                        xs.append(ops[4]); ys.append(ops[5])
                    elif op in ('v', 'y'):
                        xs.append(ops[-2]); ys.append(ops[-1])

                elif op in PATH_PAINT:
                    if xs and (max(xs) - min(xs)) < min_size and (max(ys) - min(ys)) < min_size:
                        dropped += 1
                    else:
                        out.extend(pending)
                        out.append((operands, operator))
                    pending, xs, ys = [], [], []

                else:
                    out.extend(pending); pending, xs, ys = [], [], []
                    out.append((operands, operator))

            out.extend(pending)
            page.Contents = pdf.make_stream(pikepdf.unparse_content_stream(out))

        pdf.save(dst)
    print(f"  Tiny-path filter: dropped {dropped} paths")

def remove_tiny_xobjects(src, dst, min_size=10.0):
    try:
        import pikepdf
    except ImportError:
        raise ImportError("pikepdf required for --remove-tiny-xobjects. Install: pip install pikepdf")

    removed_xobjs = 0
    removed_dos = 0

    with pikepdf.open(src) as pdf:
        for page in pdf.pages:
            res = page.get('/Resources')
            if not res or '/XObject' not in res:
                continue
            xobj_dict = res['/XObject']

            tiny = set()
            for key in list(xobj_dict.keys()):
                try:
                    xobj = xobj_dict[key]
                    bbox = xobj.get('/BBox')
                    if bbox is None:
                        continue
                    w = abs(float(bbox[2]) - float(bbox[0]))
                    h = abs(float(bbox[3]) - float(bbox[1]))
                    if max(w, h) < min_size:
                        tiny.add(str(key))
                except Exception:
                    continue

            if not tiny:
                continue

            instructions = list(pikepdf.parse_content_stream(page))
            out_instrs = []
            for operands, operator in instructions:
                if str(operator) == 'Do' and operands:
                    name = str(operands[0])
                    if name in tiny or '/' + name.lstrip('/') in tiny:
                        removed_dos += 1
                        continue
                out_instrs.append((operands, operator))

            page.Contents = pdf.make_stream(pikepdf.unparse_content_stream(out_instrs))

            for key in list(xobj_dict.keys()):
                if str(key) in tiny:
                    del xobj_dict[key]
                    removed_xobjs += 1

        pdf.save(dst)
    print(f"  Removed {removed_xobjs} tiny XObjects (<{min_size}pt), {removed_dos} Do calls")

def _run_gs(src, out, crop=None):
    x0, y0, x1, y1 = crop if crop else (0, 0, 0, 0)
    cmd = [gs_exe(), "-dBATCH", "-dNOPAUSE", "-sDEVICE=pdfwrite", "-dNoOutputFonts"]
    if crop is not None:
        w, h = x1 - x0, y1 - y0
        cmd += ["-dFIXEDMEDIA", f"-dDEVICEWIDTHPOINTS={w}", f"-dDEVICEHEIGHTPOINTS={h}"]
    cmd.append(f"-sOutputFile={out}")
    if crop is not None:
        cmd += ["-c", f"<< /PageOffset [{-x0} {-y0}] >> setpagedevice", "-f"]
    cmd.append(src)
    subprocess.run(cmd, check=True)

def flatten(src, no_images=False, crop=None, min_path_size=0.0, remove_xobjects=False):
    from pypdf import PdfReader
    src = os.path.abspath(src)
    out = os.path.splitext(src)[0] + " flat.pdf"
    reader = PdfReader(src)
    if crop is None:
        crop = _annotation_rect(reader.pages[0])
        if crop:
            print(f"  Found annotation rect: {tuple(round(v, 1) for v in crop)}")

    needs_sanitize = no_images or crop is not None
    needs_xobj     = remove_xobjects
    needs_filter   = min_path_size > 0
    tmps = []

    def mktmp():
        fd, p = tempfile.mkstemp(suffix='.pdf')
        os.close(fd); tmps.append(p); return p

    try:
        current = src
        if needs_sanitize:
            tmp_san = mktmp()
            sanitize(reader, tmp_san, no_images=no_images, crop=crop)
            current = tmp_san

        gs_out = mktmp() if (needs_xobj or needs_filter) else out
        _run_gs(current, gs_out, crop=crop if needs_sanitize else None)
        current = gs_out

        if needs_xobj:
            xobj_out = mktmp() if needs_filter else out
            remove_tiny_xobjects(current, xobj_out)
            current = xobj_out

        if needs_filter:
            filter_tiny_paths(current, out, min_path_size)
    finally:
        for t in tmps:
            if os.path.exists(t): os.unlink(t)

    print("Saved:", out)
    return out

if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Flatten SHX fonts; optionally remove images and crop")
    p.add_argument("src")
    p.add_argument("--no-images", action="store_true", help="Replace raster images with blank placeholders")
    p.add_argument("--crop", nargs=4, type=float, metavar=("X0", "Y0", "X1", "Y1"),
                   help="Crop box in PDF pts (origin bottom-left). E.g. --crop 0 0 2050 1728")
    p.add_argument("--min-path-size", type=float, default=0.0, metavar="PTS",
                   help="Drop paths whose bounding box is smaller than PTS×PTS points (default: off)")
    p.add_argument("--remove-tiny-xobjects", action="store_true",
                   help="Remove XObjects whose BBox max-dimension < 10pt (CAD hatch patterns)")
    args = p.parse_args()
    flatten(args.src, no_images=args.no_images, crop=tuple(args.crop) if args.crop else None,
            min_path_size=args.min_path_size, remove_xobjects=args.remove_tiny_xobjects)
