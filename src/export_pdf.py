"""Export slides with installed fonts and an explicit macOS DIN fallback.

PPTX retains the reference's Bahnschrift / Arial / Consolas names. Bahnschrift
is bundled with Windows and absent on this Mac. For PDF previews only, the
installed DIN Alternate (same design family) is used instead of a serif font.
No fonts are downloaded, redistributed or installed by this script.
"""
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
from xml.sax.saxutils import escape

ROOT=Path(__file__).resolve().parents[1]


def export():
    binary=shutil.which('soffice') or shutil.which('libreoffice')
    if not binary:raise RuntimeError('Install LibreOffice to export the presentation PDF.')
    with tempfile.TemporaryDirectory(prefix='stroy-export-') as tmp:
        work=Path(tmp);env=os.environ.copy()
        dirs=[Path('/System/Library/Fonts'),Path('/System/Library/Fonts/Supplemental'),Path('/Library/Fonts'),Path.home()/'Library/Fonts',Path('/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts')]
        if Path('/System/Library/Fonts/Supplemental/DIN Alternate Bold.ttf').exists():
            # Prefer the exact Bahnschrift when installed; DIN is an explicit fallback.
            entries=''.join(f'<dir>{escape(str(p))}</dir>' for p in dirs if p.exists())
            config=work/'fonts.conf'
            config.write_text('<?xml version="1.0"?><!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd"><fontconfig>'+entries+
                f'<cachedir>{escape(str(work/"font-cache"))}</cachedir>'+
                '<alias><family>Bahnschrift</family><prefer><family>Bahnschrift</family><family>DIN Alternate</family></prefer></alias></fontconfig>')
            env['FONTCONFIG_FILE']=str(config)
        subprocess.run([binary,f'-env:UserInstallation={(work/"lo-profile").as_uri()}',
                        '--headless','--convert-to','pdf','--outdir',str(ROOT/'reports'),str(ROOT/'presentation.pptx')],env=env,check=True,timeout=120)


if __name__=='__main__':export()
