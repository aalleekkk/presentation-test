import json, sys, time, urllib.request, urllib.parse, os
UA = "EduSlideDeckBuilder/1.0 (educational presentation; python-urllib)"
API = "https://commons.wikimedia.org/w/api.php"
def get(url):
    for i in range(5):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            return urllib.request.urlopen(req, timeout=60).read()
        except Exception as e:
            print("retry", e, file=sys.stderr); time.sleep(30*(i+1))
    print("fail", url, file=sys.stderr); return None
files = json.load(open(sys.argv[1]))
out = sys.argv[2]
mp=os.path.join(out,"meta.json"); meta = json.load(open(mp)) if os.path.exists(mp) else {}
for key, title in files.items():
    if any(f.startswith(key+".") for f in os.listdir(out)): continue
    if title.startswith("search:"):
        r = get(API+"?"+urllib.parse.urlencode({"action":"query","format":"json","list":"search","srnamespace":6,"srlimit":1,"srsearch":title[7:]}))
        if r is None: continue
        res = json.loads(r)["query"]["search"]
        if not res: print("NOHIT", key); continue
        title = res[0]["title"][5:]; time.sleep(6)
    q = {"action":"query","format":"json","prop":"imageinfo","iiprop":"url|size|extmetadata","iiurlwidth":"1280","titles":"File:"+title}
    r = get(API+"?"+urllib.parse.urlencode(q))
    if r is None: continue
    d = json.loads(r)
    page = next(iter(d["query"]["pages"].values()))
    if "imageinfo" not in page:
        print("MISSING", key, title); continue
    ii = page["imageinfo"][0]; em = ii.get("extmetadata", {})
    url = (ii.get("thumburl") or ii["url"]).replace("https://thumb.wikimedia.org/","https://upload.wikimedia.org/").split("?")[0]
    data = get(url)
    if data is None: continue
    ext = os.path.splitext(url)[1].lower() or ".jpg"
    open(os.path.join(out, key+ext), "wb").write(data)
    meta[key] = {"title": title, "page": ii.get("descriptionurl"), "w": ii["width"], "h": ii["height"],
      "artist": em.get("Artist",{}).get("value",""), "license": em.get("LicenseShortName",{}).get("value",""),
      "date": em.get("DateTimeOriginal",{}).get("value","")}
    json.dump(meta, open(mp,"w"), indent=1, ensure_ascii=False)
    print("OK", key, ii["width"], ii["height"], meta[key]["license"])
    time.sleep(8)
json.dump(meta, open(os.path.join(out,"meta.json"),"w"), indent=1, ensure_ascii=False)
