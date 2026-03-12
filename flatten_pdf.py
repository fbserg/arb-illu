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

def _run_gs(src, out, extra_flags=None):
    cmd = [gs_exe(), "-dBATCH", "-dNOPAUSE", "-sDEVICE=pdfwrite", "-dNoOutputFonts"]
    if extra_flags:
        cmd.extend(extra_flags)
    cmd.append(f"-sOutputFile={out}")
    cmd.append(src)
    subprocess.run(cmd, check=True)

def flatten(src, no_images=False, crop=None):
    from pypdf import PdfReader
    src = os.path.abspath(src)
    out = os.path.splitext(src)[0] + " flat.pdf"
    reader = PdfReader(src)
    if crop is None:
        crop = _annotation_rect(reader.pages[0])
        if crop:
            print(f"  Found annotation rect: {tuple(round(v, 1) for v in crop)}")
    if no_images or crop is not None:
        tmp_fd, tmp = tempfile.mkstemp(suffix='.pdf', prefix='_sanitize_')
        os.close(tmp_fd)
        try:
            sanitize(reader, tmp, no_images=no_images, crop=crop)
            _run_gs(tmp, out, extra_flags=["-dUseCropBox"] if crop is not None else None)
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)
    else:
        _run_gs(src, out)
    print("Saved:", out)
    return out

if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Flatten SHX fonts; optionally remove images and crop")
    p.add_argument("src")
    p.add_argument("--no-images", action="store_true", help="Replace raster images with blank placeholders")
    p.add_argument("--crop", nargs=4, type=float, metavar=("X0", "Y0", "X1", "Y1"),
                   help="Crop box in PDF pts (origin bottom-left). E.g. --crop 0 0 2050 1728")
    args = p.parse_args()
    flatten(args.src, no_images=args.no_images, crop=tuple(args.crop) if args.crop else None)
