import calendar
from datetime import date, timedelta
import random
from openpyxl import Workbook
from openpyxl.styles import Font
from collections import defaultdict

monthly_stats = defaultdict(lambda: defaultdict(lambda: {"bayram":0,"haftasonu":0,"normal":0}))

MAX_SAME_WEEKDAY = 2
WEEKDAY_PENALTY = 1.5
DENGE_KATSAYI = 0.7
AGIR_GUN_FRENI = 0.4
MIN_GAP_DAYS = 14

eklenme_tarihi = {}
cikma_tarihi = {}

# GECMIS_YUK aynen bırakıldı
GECMIS_YUK = {...}   # (senin verdiğin uzun veri aynen burada kalacak)

def turkiye_tatilleri(year):
    return {
        date(year,1,1), date(year,4,23), date(year,5,1),
        date(year,5,19), date(year,7,15),
        date(year,8,30), date(year,10,29),
        date(2026,3,20),date(2026,3,21),date(2026,3,22),
        date(2026,5,27),date(2026,5,28),
        date(2026,5,29),date(2026,5,30),
    }

def arefe_gunleri(year):
    return {date(2026,3,19), date(2026,5,26)}

def day_weight(d, tatil,arefe):
    if d in tatil or d.weekday()==6: return 2.0
    if d.weekday()==5 or d in arefe: return 1.5
    return 1.0

def score_person(p,d,w,totals,counts,weekday_stats,last_dates):

    skor = totals[p]*DENGE_KATSAYI + counts[p]

    if weekday_stats[p][d.weekday()] >= MAX_SAME_WEEKDAY:
        skor += WEEKDAY_PENALTY

    if p in last_dates:
        gap = (d-last_dates[p]).days
        skor -= min(gap,30)*0.05

    if w>1.4:
        skor += AGIR_GUN_FRENI * weekday_stats[p][d.weekday()]

    return skor + random.random()*0.01

def zorunlu_secim(grup,d,w,tatil,totals,counts,weekday_stats,last_dates,bayram_stats):

    kademe1, kademe2, kademe3 = [], [], []

    for p in grup:

        if p in eklenme_tarihi and d < eklenme_tarihi[p]:
            continue

        if p in cikma_tarihi and d >= cikma_tarihi[p]:
            continue

        gap = (d-last_dates[p]).days if p in last_dates else 999

        if gap >= MIN_GAP_DAYS and weekday_stats[p][d.weekday()] < MAX_SAME_WEEKDAY:
            kademe1.append(p)
        elif weekday_stats[p][d.weekday()] < MAX_SAME_WEEKDAY:
            kademe2.append(p)
        else:
            kademe3.append(p)

    adaylar = kademe1 or kademe2 or kademe3 or grup

    if d in tatil:
        min_b = min(bayram_stats[p] for p in adaylar)
        adaylar = [p for p in adaylar if bayram_stats[p]==min_b]

    return min(adaylar, key=lambda p: score_person(p,d,w,totals,counts,weekday_stats,last_dates))


# create_groups() SENDEKİYLE AYNI (dokunmadım)

def create_groups():
    return { ... }  # grup listelerin aynen burada kalacak


KOMB_ABC=[("A1","B2","C3"),("B1","C2","A3"),("C1","A2","B3")]
KOMB_DEG=[("D1","E2","G3"),("E1","G2","D3"),("G1","D2","E3")]
F_ROTASYON=["F1","F2","F3"]


def generate_month(groups, year, month, totals, counts, weekday_stats, bayram_stats, last_dates):

    tatil = turkiye_tatilleri(year)
    arefe = arefe_gunleri(year)

    first = date(year, month, 1)
    dim = calendar.monthrange(year, month)[1]

    schedule = {}

    for i in range(dim):

        d = first + timedelta(days=i)
        w = day_weight(d, tatil,arefe)

        picks = {}

        for g in KOMB_ABC[i % 3] + KOMB_DEG[i % 3]:

            pick = zorunlu_secim(groups[g], d, w, tatil, totals, counts, weekday_stats, last_dates, bayram_stats)

            picks[g] = pick

            totals[pick] += w
            counts[pick] += 1
            weekday_stats[pick][d.weekday()] += 1
            last_dates[pick] = d

            if d in tatil:
                bayram_stats[pick] += 1

            # 🔵 AYLIK ISTATISTIK
            key = (d.year, d.month)

            if d in tatil:
                monthly_stats[pick][key]["bayram"] += 1
            elif d.weekday() >= 5:
                monthly_stats[pick][key]["haftasonu"] += 1
            else:
                monthly_stats[pick][key]["normal"] += 1

        fg = F_ROTASYON[i % 3]

        pick = zorunlu_secim(groups[fg], d, w, tatil, totals, counts, weekday_stats, last_dates, bayram_stats)

        picks[fg] = pick

        totals[pick] += w
        counts[pick] += 1
        weekday_stats[pick][d.weekday()] += 1
        last_dates[pick] = d

        # 🔵 AYLIK ISTATISTIK
        key = (d.year, d.month)

        if d in tatil:
            monthly_stats[pick][key]["bayram"] += 1
        elif d.weekday() >= 5:
            monthly_stats[pick][key]["haftasonu"] += 1
        else:
            monthly_stats[pick][key]["normal"] += 1

        schedule[d] = picks

    return schedule


def main(y,m,nm):

    groups=create_groups()

    totals={p:0 for g in groups.values() for p in g}
    counts={p:0 for g in groups.values() for p in g}
    weekday_stats={p:{i:0 for i in range(7)} for p in totals}
    bayram_stats={p:0 for p in totals}
    last_dates={}
    gecmis_katsayi_map={}

    for p,v in GECMIS_YUK.items():

        if p not in totals:
            continue

        kats = v["normal"] + v["haftasonu"]*1.5 + v["bayram"]*2

        totals[p]+=kats
        counts[p]+=v["normal"]+v["haftasonu"]+v["bayram"]
        bayram_stats[p]+=v["bayram"]

        weekday_stats[p][5]+=v["haftasonu"]//2
        weekday_stats[p][6]+=v["haftasonu"]//2

        gecmis_katsayi_map[p]=round(kats,2)

    for p in totals:
        gecmis_katsayi_map.setdefault(p,0)

    wb=Workbook()

    gun=["Pzt","Salı","Çarş","Perş","Cuma","Ctesi","Pazar"]
    header=["Tarih","Gün"]+list(groups.keys())

    for k in range(nm):

        year=y+((m-1+k)//12)
        month=((m-1+k)%12)+1

        ws=wb.create_sheet(f"{year}-{month:02d}")

        ws.append(header)

        sched=generate_month(groups,year,month,totals,counts,weekday_stats,bayram_stats,last_dates)

        for d,p in sorted(sched.items()):

            row=[d.strftime("%d.%m.%Y"),gun[d.weekday()]]

            for g in groups:
                row.append(p.get(g,""))

            ws.append(row)

        for c in ws[1]:
            c.font=Font(bold=True)

    summary=wb.create_sheet("GENEL OZET")

    summary.append([
        "Eczane","Grup","Geçmiş Katsayı","Geçmiş Bayram",
        "Toplam Nöbet","Toplam Katsayı","Bayram",
        "Pzt","Salı","Çarş","Perş","Cuma","Ctesi","Pazar"
    ])

    eczane_grup={p:g for g,plist in groups.items() for p in plist}

    for p in totals:

        summary.append([
            p,
            eczane_grup.get(p,""),
            gecmis_katsayi_map.get(p,0),
            GECMIS_YUK.get(p,{}).get("bayram",0),
            counts[p],
            round(totals[p],2),
            bayram_stats[p],
            weekday_stats[p][0],
            weekday_stats[p][1],
            weekday_stats[p][2],
            weekday_stats[p][3],
            weekday_stats[p][4],
            weekday_stats[p][5],
            weekday_stats[p][6],
        ])

    for c in summary[1]:
        c.font = Font(bold=True)

    wb.remove(wb["Sheet"])
    wb.save("Son.xlsx")

    # 🔵 AYLIK DETAY EXCEL

    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "AYLIK DETAY"
    ws2.append(["Eczane","Yıl","Ay","Bayram","Hafta Sonu","Normal"])

    for eczane in sorted(monthly_stats.keys()):
        for (yil, ay), veri in sorted(monthly_stats[eczane].items()):
            ws2.append([
                eczane,
                yil,
                ay,
                veri["bayram"],
                veri["haftasonu"],
                veri["normal"]
            ])

    for c in ws2[1]:
        c.font = Font(bold=True)

    wb2.save("aylik_nobet_data.xlsx")

    return "Son.xlsx","aylik_nobet_data.xlsx"


def run_schedule(y,m,nm):
    return main(y,m,nm)
