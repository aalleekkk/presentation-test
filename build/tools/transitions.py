"""Inject slide transitions (Morph with Fade fallback, fade, fade-through-black) into a pptx."""
import json, re, sys, zipfile, shutil, os
src, trans_json = sys.argv[1], sys.argv[2]
T = {int(k): v for k, v in json.load(open(trans_json)).items()}
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"
def xml_for(kind):
    if kind == "morph":
        return (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:p159="{P159}" Requires="p159">'
                f'<p:transition xmlns:p14="{P14}" spd="slow" p14:dur="1600"><p159:morph option="byObject"/></p:transition>'
                f'</mc:Choice><mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>')
    if kind == "fadeBlack":
        return '<p:transition spd="slow"><p:fade thruBlk="1"/></p:transition>'
    return '<p:transition spd="med"><p:fade/></p:transition>'
tmp = src + ".tmp"
with zipfile.ZipFile(src) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", item.filename)
        if m and int(m.group(1)) in T:
            x = data.decode("utf8")
            x = re.sub(r"<p:transition.*?</p:transition>|<mc:AlternateContent.*?</mc:AlternateContent>", "", x, flags=re.S)
            tx = xml_for(T[int(m.group(1))])
            if "</p:clrMapOvr>" in x: x = x.replace("</p:clrMapOvr>", "</p:clrMapOvr>" + tx, 1)
            elif "<p:timing" in x: x = x.replace("<p:timing", tx + "<p:timing", 1)
            else: x = x.replace("</p:sld>", tx + "</p:sld>")
            data = x.encode("utf8")
        zout.writestr(item, data)
shutil.move(tmp, src)
print("transitions:", T)
