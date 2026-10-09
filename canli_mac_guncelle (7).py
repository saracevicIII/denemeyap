#!/usr/bin/env python3
"""
Canlı maç yayınlarını kendi kanal listene ekler (birincil kaynak + yedek kaynaklar).

Liste olarak Excel (kanal_listesi_guncel_epg_logo.xlsx, "Tüm Liste" sayfası) ya da .m3u verebilirsin.
Excel verirsen ana liste Excel'dir; script tüm m3u'yu ondan üretir, canlı maçları en başa koyar
ve canlı kanalların logo / tvg-id bilgisini listendeki aynı isimli kanaldan alır.

Kullanım:
  python canli_mac_guncelle.py --liste kanal_listesi_guncel_epg_logo.xlsx
  python canli_mac_guncelle.py --liste kanal.xlsx --yedek https://yedek-site/sayfa --yedek baska.m3u
  python canli_mac_guncelle.py --liste kanal.xlsx --html kayitli_platinsport.html   (test)

Saatler: tüm canlı yayınların saati Türkiye saatine (UTC+3) çevrilir. Platinsport UTC verdiği için
doğrudan çevrilir; yedek kaynakta sadece "19:00" gibi saat varsa dilimi --yedek-tz ile söyle (varsayılan Europe/Madrid).

Kaynak sırası: 1) Platinsport  2) --yedek ile verdiklerin (verdiğin sırayla).
İlk çalışan kaynak kullanılır. Hiçbiri çalışmazsa çıktı DOSYASI DEĞİŞTİRİLMEZ.
--yedek şunlardan biri olabilir: web sayfası (HTML), .m3u adresi/dosyası, kayıtlı HTML dosyası.

Kurulum:
  pip install beautifulsoup4 lxml openpyxl playwright
  playwright install chromium
"""
import argparse, html, os, re, sys, time, urllib.request
from datetime import datetime, timezone, timedelta
try:
    from bs4 import BeautifulSoup
except ImportError:
    sys.exit("HATA: 'beautifulsoup4' kurulu değil. Pydroid 3: menü > Pip > 'beautifulsoup4' yaz > Install.\n"
             "Bilgisayarda: pip install beautifulsoup4")
try:
    import lxml  # noqa: F401
    PARSER = "lxml"
except ImportError:
    PARSER = "html.parser"          # lxml yoksa (ör. telefon) yerleşik ayrıştırıcı kullanılır
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:                   # Pydroid exec() ile çalıştırınca __file__ olmayabilir
    BASE_DIR = os.getcwd()

# ---- AYARLAR (telefonda / Pydroid'de komut yazamadığın için buradan değiştir) ----
LISTE = "kanal_listesi_guncel_epg_logo.xlsx"   # Excel veya .m3u; tam yol da yazabilirsin: "/storage/emulated/0/Download/liste.xlsx"
CIKTI = "kanal_listesi_canli.m3u"              # çıktı, liste dosyasıyla aynı klasöre yazılır
SITE_ALANLARI = ["platinmax.com", "platinsport.com"]   # sırayla denenir; biri açılmazsa diğerine geçer
HTML_DOSYA = ""                                # tarayıcıdan kaydettiğin PLAY sayfasının (HTML) yolu; boş bırakırsan siteden çeker
YEDEKLER = []                                  # örn. ["https://ornek.com/sayfa", "/storage/emulated/0/Download/yedek.m3u"]
# --------------------------------------------------------------------------------
GRUP = "Canlı Maç"
MIN_YAYIN = 5                               # bundan az yayın bulan kaynak "çalışmıyor" sayılır
TR = timezone(timedelta(hours=3))           # Türkiye: UTC+3
ACE = "http://127.0.0.1:6878/ace/getstream?id="
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
LIG_ANAHTAR = ['league','liga','serie','bundesliga','ligue','eredivisie','championship','cup','portugal',
               'primeira','super','pro league','paulista','carioca','profesional','division','nations',
               'fifa','uefa','nba','formula','motogp','tennis']

# ------------------------------------------------------------------ yardımcılar
ARAMA_KLASORLERI = ["/storage/emulated/0/Download", "/storage/emulated/0/Documents", "/storage/emulated/0",
                    "/sdcard/Download", "/sdcard/Documents", "/sdcard", "~/Download", "~/Downloads"]

def yol_bul(yol):
    """Dosyayı sırayla şurada arar: verdiğin yol, bulunduğun klasör, scriptin klasörü, telefonun
    İndirilenler/Belgeler klasörleri. Hâlâ yoksa adında 'kanal_listesi' geçen en yeni aynı türdeki dosyayı alır."""
    if os.path.isabs(yol) and os.path.exists(yol):
        return yol
    klasorler = [os.getcwd(), BASE_DIR] + [os.path.expanduser(k) for k in ARAMA_KLASORLERI]
    for k in klasorler:
        aday = os.path.join(k, yol)
        if os.path.exists(aday):
            return aday
    import glob
    uzanti = os.path.splitext(yol)[1] or ".*"
    adaylar = []
    for k in klasorler:
        adaylar += glob.glob(os.path.join(k, "*kanal_listesi*" + uzanti))
    if adaylar:
        en_yeni = max(adaylar, key=os.path.getmtime)
        print(f"ℹ '{yol}' bulunamadı, en yeni benzer dosya kullanılıyor: {en_yeni}")
        return en_yeni
    return yol

def sayfa_oku(yol):
    """HTML ya da tarayıcının 'İndir' ile kaydettiği .mhtml dosyasını metin olarak okur."""
    veri = open(yol, "rb").read()
    if b"MIME-Version" in veri[:3000] or veri[:30].lower().startswith((b"from:", b"subject:", b"date:")):
        import email, email.policy
        msg = email.message_from_bytes(veri, policy=email.policy.default)
        en_iyi = ""
        for parca in msg.walk():
            if parca.get_content_type() == "text/html":
                try:
                    metin = parca.get_content()
                except Exception:
                    metin = parca.get_payload(decode=True).decode("utf-8", "replace")
                if "acestream://" in metin:
                    return metin
                en_iyi = en_iyi or metin
        return en_iyi
    return veri.decode("utf-8", errors="replace")

def kayitli_sayfa_bul(saat=18):
    """İndirilenler vb. klasörlerde adında 'platin' geçen, son <saat> saatte kaydedilmiş en yeni sayfayı bulur."""
    import glob
    klasorler = [os.getcwd(), BASE_DIR] + [os.path.expanduser(k) for k in ARAMA_KLASORLERI]
    adaylar = []
    for k in klasorler:
        for kalip in ("*latin*.mhtml", "*latin*.mht", "*latin*.html", "*latin*.htm"):
            adaylar += glob.glob(os.path.join(k, kalip))
    adaylar = [f for f in set(adaylar) if time.time() - os.path.getmtime(f) < saat * 3600]
    return max(adaylar, key=os.path.getmtime) if adaylar else None

def temiz(s):
    s = html.unescape(s or "").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()

def kanal_adi_temizle(ad, kod="XX"):
    ad = re.sub(r"\b(STREAM|4K|FHD|UHD)\b", "", temiz(ad), flags=re.I).strip()
    return f"Yayın {kod}" if ad.upper() in ("", "HD") else ad

def norm(ad):
    ad = re.sub(r"\[[^\]]*\]|\([^)]*\)", "", ad.lower())
    ad = re.sub(r"\b(hd|fhd|uhd|4k|1080p|720p)\b", "", ad)
    return re.sub(r"[^a-z0-9а-я]", "", ad)

def bayrak(a):
    f = a.find("span", class_=re.compile(r"\bfi\b|\bfi-"))
    if f:
        for c in f.get("class", []):
            if c.startswith("fi-") and len(c) == 5:
                return c[3:].upper().replace("UK", "GB")
    return "XX"

def tr_saat(dt_str):
    try:
        return datetime.fromisoformat(dt_str.replace("Z", "+00:00")).astimezone(TR)
    except Exception:
        return None

AVRUPA_ORTA = {"Europe/Madrid", "Europe/Paris", "Europe/Berlin", "Europe/Rome", "Europe/Amsterdam",
                "Europe/Brussels", "Europe/Vienna", "Europe/Zurich", "Europe/Warsaw", "Europe/Prague"}

def saat_cevir(hhmm, tz_adi):
    """'19:00' (tz_adi saat diliminde, bugün) -> Türkiye saati 'HH:MM'. Çevrilemezse aynen döner."""
    s, d = map(int, hhmm.split(":"))
    try:
        from zoneinfo import ZoneInfo
        z = ZoneInfo(tz_adi)
        bugun = datetime.now(z).replace(hour=s, minute=d, second=0, microsecond=0)
        return bugun.astimezone(TR).strftime("%H:%M")
    except Exception:
        pass
    if tz_adi in AVRUPA_ORTA:           # saat dilimi veritabanı yoksa: Orta Avrupa kuralı (UTC+1, yazın +2)
        simdi = datetime.now(timezone.utc)
        def son_pazar(ay):
            g = datetime(simdi.year, ay, 31, 1, 0, tzinfo=timezone.utc)
            while g.weekday() != 6:
                g -= timedelta(days=1)
            return g
        yaz = son_pazar(3) <= simdi < son_pazar(10)
        fark = 2 if yaz else 1
        return f"{(s - fark + 3) % 24:02d}:{d:02d}"
    print(f"⚠ Saat çevrilemedi ({hhmm}, {tz_adi}) – olduğu gibi bırakıldı")
    return hhmm

import http.cookiejar
_CEREZ = http.cookiejar.CookieJar()
_ACICI = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(_CEREZ))

def indir_son(adres, referer=None):
    """(metin, son_adres) döndürür; çerezleri saklar, yönlendirmeleri izler."""
    basliklar = {"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9", "Cookie": "disclaimer_accepted=true"}
    if referer:
        basliklar["Referer"] = referer
    with _ACICI.open(urllib.request.Request(adres, headers=basliklar), timeout=30) as r:
        return r.read().decode("utf-8", errors="replace"), r.geturl()

def indir(adres, referer=None):
    return indir_son(adres, referer)[0]

# ------------------------------------------------------------------ KAYNAK 1: Platinsport
def platinsport_cek(alan):
    from playwright.sync_api import sync_playwright
    yakalanan = {}
    with sync_playwright() as p:
        tarayici = p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        ctx = tarayici.new_context(user_agent=UA, viewport={"width": 1920, "height": 1080},
                                   locale="en-US", ignore_https_errors=True)
        ctx.add_cookies([{"name": "disclaimer_accepted", "value": "true", "domain": "." + alan,
                          "path": "/", "expires": int((datetime.now(timezone.utc) + timedelta(days=1)).timestamp())}])
        def yol(route, request):
            if "source-list.php" in request.url:
                yanit = route.fetch()
                yakalanan["html"] = yanit.text()
                route.fulfill(response=yanit)
            else:
                route.continue_()
        ctx.route("**/*", yol)
        sayfa = ctx.new_page()
        for adres in (f"https://www.{alan}/", f"https://{alan}/"):
            try:
                sayfa.goto(adres, timeout=60000, wait_until="domcontentloaded")
                break
            except Exception as e:
                hata = e
        else:
            raise hata
        time.sleep(2)
        buton = sayfa.locator('a[href*="source-list.php"]').first      # platinmax: PLAY bağlantısı
        buton.wait_for(state="visible", timeout=20000)
        buton.click()                                                   # popup ya da aynı sekme olabilir
        for _ in range(40):                                             # yanıt yakalanana kadar en çok ~20 sn bekle
            if "html" in yakalanan:
                break
            time.sleep(0.5)
        tarayici.close()
    if "html" not in yakalanan:
        raise RuntimeError("source-list.php yanıtı yakalanamadı")
    return yakalanan["html"]

def platinsport_duz_cek(alan):
    """Tarayıcı (playwright) olmadan (ör. telefon): ana sayfadaki PLAY / günlük sayfa linkini
    ana sayfa referer'i ve çerezle doğrudan indirir. Site tarayıcı doğrulaması isterse (ana sayfaya
    geri yönlendirirse) bu yol çalışmaz, o zaman 'tarayıcı gerekiyor' hatası verir."""
    from urllib.parse import urljoin
    hata = None
    for adres in (f"https://www.{alan}/", f"https://{alan}/"):
        try:
            ana, ana_son = indir_son(adres)
        except Exception as e:
            hata = e
            continue
        adaylar = []
        for href in re.findall(r'href=["\']([^"\']+)["\']', ana):
            if re.search(r"source-list\.php", href, re.I):
                adaylar.append(href)
            else:
                m = re.search(r"((?:https?:)?[^\s\"']*?/link/\d{2}[a-z]{3}[a-z0-9]+/01\.php)", href, re.I)
                if m:
                    g = m.group(1)
                    adaylar.append(g[g.rfind("http"):] if g.rfind("http") > 0 else g)   # bc.vc gibi sarmalı çöz
        if not adaylar:
            raise RuntimeError("ana sayfada PLAY / günlük sayfa linki bulunamadı")
        gorulen = set()
        for href in adaylar:
            hedef = urljoin(ana_son, href)
            if hedef in gorulen:
                continue
            gorulen.add(hedef)
            sayfa, son = indir_son(hedef, referer=ana_son)
            sonuc = platinsport_coz(sayfa)
            if len(sonuc) < MIN_YAYIN:
                sonuc = genel_html_coz(sayfa)
            if len(sonuc) >= MIN_YAYIN:
                return sonuc
        raise RuntimeError("PLAY sayfası yayın listesi vermedi (ana sayfaya yönlendirilmiş olabilir: site tarayıcı doğrulaması istiyor)")
    raise hata

def platinsport_coz(icerik):
    soup = BeautifulSoup(icerik, PARSER)
    sonuc, lig = [], "Diğer"
    for el in soup.find_all(["p", "div"]):
        if el.name == "p":
            t = el.get_text().strip()
            if t and any(k in t.lower() for k in LIG_ANAHTAR):
                lig = t
        elif "match-title-bar" in (el.get("class") or []):
            zaman = el.find("time")
            dt = tr_saat(zaman.get("datetime", "")) if zaman else None
            kopya = BeautifulSoup(str(el), PARSER).find("div")
            for z in kopya.find_all("time"):
                z.decompose()
            mac = temiz(kopya.get_text())
            grup = el.find_next_sibling("div", class_="button-group")
            if not grup:
                continue
            for a in grup.find_all("a", href=re.compile(r"^acestream://")):
                kod = bayrak(a)
                k = BeautifulSoup(str(a), PARSER).find("a")
                for f in k.find_all("span", class_=re.compile(r"\bfi\b|\bfi-")):
                    f.decompose()
                kanal = kanal_adi_temizle(k.get_text() or a.get("title", ""), kod)
                sonuc.append(dict(dt=dt, saat=None, lig=lig, mac=mac, kanal=kanal, ulke=kod,
                                  url=ACE + a["href"].replace("acestream://", "").strip()))
    return sonuc

# ------------------------------------------------------------------ YEDEK kaynaklar
def genel_html_coz(icerik, tz_adi="Europe/Madrid"):
    """Platinsport düzenine benzemeyen sayfalar için: sayfadaki her acestream:// bağlantısını alır,
    maç adını bağlantıdan önceki en yakın 'X vs Y' / 'X - Y' metninden tahmin eder."""
    soup = BeautifulSoup(icerik, PARSER)
    sonuc = []
    for a in soup.find_all("a", href=re.compile(r"^acestream://", re.I)):
        kod = bayrak(a)
        mac, saat = "Etkinlik", None
        onceki = a.find_previous(string=re.compile(r"\bvs\.?\b|\s-\s|\bv\b", re.I))
        if onceki:
            t = temiz(onceki)
            if 4 < len(t) < 120:
                mac = t
                m = re.search(r"\b(\d{1,2}:\d{2})\b", t)
                if m:
                    saat = saat_cevir(m.group(1), tz_adi)
                    mac = temiz(t.replace(m.group(1), ""))
        kanal = kanal_adi_temizle(a.get_text() or a.get("title", ""), kod)
        sonuc.append(dict(dt=None, saat=saat, lig="", mac=mac, kanal=kanal, ulke=kod,
                          url=ACE + re.sub(r"^acestream://", "", a["href"], flags=re.I).strip()))
    return sonuc

def m3u_coz(icerik, tz_adi="Europe/Madrid"):
    """Başka birinin hazır m3u'su: başlıklar olduğu gibi alınır (saatleri DÖNÜŞTÜRÜLMEZ)."""
    satirlar = icerik.replace("\r\n", "\n").split("\n")
    sonuc, i = [], 0
    while i < len(satirlar) - 1:
        if satirlar[i].startswith("#EXTINF"):
            ad = satirlar[i].rsplit(",", 1)[-1].strip()
            tvg = re.search(r'tvg-name="([^"]*)"', satirlar[i])
            url = satirlar[i + 1].strip()
            ad = re.sub(r"^(\d{1,2}:\d{2})", lambda m: saat_cevir(m.group(1), tz_adi), ad)
            if url and not url.startswith("#"):
                sonuc.append(dict(dt=None, saat=None, lig="", mac="", kanal=ad, ulke="XX", url=url, ham=True,
                                  anahtar=tvg.group(1) if tvg else ad))
            i += 2
        else:
            i += 1
    return sonuc

def yedek_coz(kaynak, tz_adi="Europe/Madrid"):
    icerik = open(kaynak, encoding="utf-8").read() if os.path.exists(kaynak) else indir(kaynak)
    if icerik.lstrip().startswith("#EXTM3U"):
        return m3u_coz(icerik, tz_adi)
    sonuc = platinsport_coz(icerik)
    return sonuc if len(sonuc) >= MIN_YAYIN else genel_html_coz(icerik, tz_adi)

def kaynaklardan_cek(kaynaklar):
    """kaynaklar: [(ad, fonksiyon), ...] sırayla denenir, ilk çalışan döner."""
    for ad, fonk in kaynaklar:
        try:
            girdiler = fonk()
            if len(girdiler) >= MIN_YAYIN:
                print(f"✓ Kaynak: {ad} ({len(girdiler)} yayın)")
                return girdiler
            print(f"✗ {ad}: yalnızca {len(girdiler)} yayın bulundu, sonrakine geçiliyor")
        except Exception as e:
            print(f"✗ {ad}: çalışmadı ({type(e).__name__}: {e}), sonrakine geçiliyor")
    return None

# ------------------------------------------------------------------ Liste okuma / yazma
def excel_oku(yol):
    try:
        import openpyxl
    except ImportError:
        sys.exit("HATA: Excel için 'openpyxl' gerekli. Pydroid 3: menü > Pip > 'openpyxl' > Install. Ya da listeyi .m3u ver.")
    ws = openpyxl.load_workbook(yol, data_only=True)["Tüm Liste"]
    satir = ws.iter_rows(values_only=True)
    baslik = {str(b).strip(): i for i, b in enumerate(next(satir)) if b}
    s_ad, s_grup, s_url = baslik["Kanal Adı"], baslik["Grup"], baslik["Adres"]
    s_id, s_logo = baslik.get("tvg-id"), baslik.get("Logo")
    kanallar = []
    for r in satir:
        if r[s_ad] and r[s_url]:
            kanallar.append(dict(ad=str(r[s_ad]).strip(), grup=r[s_grup] or "Diğer", url=str(r[s_url]).strip(),
                                 id=(r[s_id] or "") if s_id is not None else "",
                                 logo=(r[s_logo] or "") if s_logo is not None else ""))
    return kanallar

def m3u_oku(yol):
    satirlar = open(yol, encoding="utf-8").read().replace("\r\n", "\n").split("\n")
    kanallar, i = [], 1
    while i < len(satirlar) - 1:
        if satirlar[i].startswith("#EXTINF"):
            inf = satirlar[i]
            nitelik = lambda n: (re.search(n + r'="([^"]*)"', inf) or [None, ""])[1]
            kanallar.append(dict(ad=inf.rsplit(",", 1)[-1].strip(), grup=nitelik("group-title"),
                                 url=satirlar[i + 1].strip(), id=nitelik("tvg-id"), logo=nitelik("tvg-logo")))
            i += 2
        else:
            i += 1
    return [k for k in kanallar if k["grup"] != GRUP]       # eski canlı maç bloğunu at

def numarala(girdiler):
    gruplar = {}
    for g in girdiler:
        gruplar.setdefault((g["mac"], g["kanal"].lower()), []).append(g)
    for liste in gruplar.values():
        if len(liste) < 2:
            continue
        sayili = bool(re.search(r"\d$", liste[0]["kanal"]))
        for j, g in enumerate(liste, 1):
            if sayili:
                if j > 1:
                    g["kanal"] += f" ({j})"
            else:
                g["kanal"] += f" {j}"

def inf_satiri(ad, grup, kimlik="", logo="", tvg_ad=None):
    ozellik = []
    if kimlik: ozellik.append(f'tvg-id="{kimlik}"')
    ozellik.append(f'tvg-name="{tvg_ad or ad}"')
    if logo: ozellik.append(f'tvg-logo="{logo}"')
    ozellik.append(f'group-title="{grup}"')
    return f'#EXTINF:-1 {" ".join(ozellik)},{ad}'

def canli_satirlari(girdiler, kanallar):
    sozluk = {}
    for k in kanallar:                      # aynı isim varsa ilk (üstteki) kanal geçerli
        sozluk.setdefault(norm(k["ad"]), k)
    maxdt = datetime.max.replace(tzinfo=TR)
    girdiler.sort(key=lambda g: (g["dt"] or maxdt, g["saat"] or "", g["lig"], g["mac"]))
    numarala([g for g in girdiler if not g.get("ham")])
    out, eslesen = [], 0
    for g in girdiler:
        if g.get("ham"):
            k = sozluk.get(norm(g["anahtar"]))
            eslesen += bool(k)
            out += [inf_satiri(g["kanal"], GRUP, k["id"] if k else "", k["logo"] if k else "", g["anahtar"]), g["url"]]
            continue
        saat = g["dt"].strftime("%H:%M") if g["dt"] else (g["saat"] or "??:??")
        parcalar = [saat] + ([g["lig"]] if g["lig"] else []) + [g["mac"], g["kanal"]]
        ad = " | ".join(p for p in parcalar if p) + (f' [{g["ulke"]}]' if g["ulke"] != "XX" else "")
        k = sozluk.get(norm(re.sub(r"\s\d+$|\s\(\d+\)$", "", g["kanal"]))) or sozluk.get(norm(g["kanal"]))
        eslesen += bool(k)
        out += [inf_satiri(ad, GRUP, k["id"] if k else "", k["logo"] if k else "", g["kanal"]), g["url"]]
    return out, eslesen

def yaz(cikti, canli, kanallar):
    satirlar = ["#EXTM3U"] + canli
    for k in kanallar:
        satirlar += [inf_satiri(k["ad"], k["grup"], k["id"], k["logo"]), k["url"]]
    with open(cikti, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(satirlar) + "\n")

# ------------------------------------------------------------------ ana akış
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--liste", default=LISTE, help=".xlsx (Tüm Liste sayfası) veya .m3u")
    ap.add_argument("--cikti", default=CIKTI)
    ap.add_argument("--yedek", action="append", default=list(YEDEKLER), help="yedek kaynak: sayfa adresi, .m3u adresi veya dosya (birden çok verilebilir)")
    ap.add_argument("--yedek-tz", default="Europe/Madrid",
                    help="yedek kaynaktaki 'HH:MM' saatlerinin dilimi (Türkiye saatine çevrilir). Platinsport bu ayardan etkilenmez, o zaten UTC verir")
    ap.add_argument("--html", default=(HTML_DOSYA or None), help="Platinsport yerine kayıtlı HTML dosyası kullan (test)")
    a = ap.parse_args()

    a.liste = yol_bul(a.liste)
    if not os.path.exists(a.liste):
        sys.exit(f"HATA: liste dosyası bulunamadı: {a.liste}\n"
                 "Dosyayı telefonun İndirilenler (Download) klasörüne koy ya da scriptin en üstündeki LISTE satırına tam yolunu yaz.")
    if not os.path.isabs(a.cikti):
        a.cikti = os.path.join(os.path.dirname(a.liste), a.cikti)
    kanallar = excel_oku(a.liste) if a.liste.lower().endswith((".xlsx", ".xlsm")) else m3u_oku(a.liste)
    kaynaklar = []
    html_yol = yol_bul(a.html) if a.html else kayitli_sayfa_bul()
    if html_yol:
        if not a.html:
            print(f"ℹ Kaydettiğin sayfa bulundu: {html_yol}")
        def kayitli(yol=html_yol):
            t = sayfa_oku(yol)
            r = platinsport_coz(t)
            return r if len(r) >= MIN_YAYIN else genel_html_coz(t)
        kaynaklar.append((f"Kayıtlı sayfa: {os.path.basename(html_yol)}", kayitli))
    try:
        import playwright.sync_api  # noqa: F401
        tarayici_var = True
    except ImportError:
        tarayici_var = False
        print("ℹ playwright yok (telefonda normal), tarayıcısız yol kullanılacak")
    for alan in SITE_ALANLARI:
        if tarayici_var:
            kaynaklar.append((f"{alan} (tarayıcı)", lambda alan=alan: platinsport_coz(platinsport_cek(alan))))
        kaynaklar.append((f"{alan} (tarayıcısız)", lambda alan=alan: platinsport_duz_cek(alan)))
    kaynaklar += [(f"Yedek {i}: {y[:60]}", lambda y=y: yedek_coz(y, a.yedek_tz))
                                                for i, y in enumerate(a.yedek, 1)]
    girdiler = kaynaklardan_cek(kaynaklar)
    if not girdiler:
        sys.exit("HATA: hiçbir kaynak çalışmadı, çıktı dosyası değiştirilmedi.\n"
                 "Çözüm: telefon tarayıcında platinmax.com'u aç, bir maçın PLAY'ine bas, açılan yayın listesi sayfasını\n"
                 "'İndir' ile kaydet (İndirilenler klasörüne .mhtml/.html olarak iner; adında 'Platin' geçmeli), sonra scripti tekrar çalıştır.")
    canli, eslesen = canli_satirlari(girdiler, kanallar)
    yaz(a.cikti, canli, kanallar)
    print(f"{len(canli)//2} canlı yayın eklendi ({eslesen} tanesine listendeki logo/tvg-id verildi), "
          f"{len(kanallar)} kanal korundu -> {a.cikti}")

if __name__ == "__main__":
    main()
