# 7 хоногийн үнийн дашбоард (ҮСХ 1212.mn)

Хүнсний гол нэр болон бензин түлшний **7 хоногийн дундаж үнэ**-ийг Улаанбаатар болон 21 аймгаар харуулах интерактив дашбоард.

**Амьд сайт:** https://batsukh1111.github.io/mongolia-weekly-price-dashboard/

## Онцлог

- **7 хоногийн тойм** — нээхэд шууд уншигдах товч мэдээ (хэд өссөн, хэд буурсан, хамгийн их хөдөлсөн нэр)
- **Хайлт** — мах, гурил, бензин гэх мэтээр шууд олох
- **Ерөнхий** — KPI, хамгийн их өөрчлөлт, УБ үнийн түвшин, нийтлэг бүтээгдэхүүнээр аймаг vs УБ
- **Газрын зураг** — 21 аймаг + УБ (Leaflet)
- **Улаанбаатар** — 30+ бүтээгдэхүүн, ангилал, түүх, олон бүтээгдэхүүн
- **Аймгууд** — 11 бүтээгдэхүүн × 21 аймаг, heatmap
- **Харьцуулалт / чиг хандлага**
- **Excel / CSV таталт** — нэвтрэхгүйгээр
- Таб бүр хуваалцах хаягтай (`#map`, `#ub`, …)

## Өгөгдлийн эх

| Бүс | Хүснэгт | API |
|-----|---------|-----|
| **Улаанбаатар** | `DT_NSO_0600_001V4` | [table-view](https://www.1212.mn/mn/statcate/table-view/Economy,%20environment/Consumer%20Price%20Index/DT_NSO_0600_001V4.px) |
| **Аймгууд** | `DT_NSO_0300_010V5` | [table-view](https://www.1212.mn/mn/statcate/table-view/Economy,%20environment/Consumer%20Price%20Index/DT_NSO_0300_010V5.px) |

- API: [https://data.1212.mn/](https://data.1212.mn/) (PxWeb JSON-stat2)
- Нэгж: төгрөг (₮)
- Давтамж: 7 хоног

## Локаль ажиллуулах

Node.js (LTS) хэрэгтэй.

```bash
cd mongolia-weekly-price-dashboard
npm install
npm run dev
```

Дараа нь: http://localhost:5173/mongolia-weekly-price-dashboard/

```bash
npm run build    # production bundle → dist/
npm run preview  # built файлыг шалгах
```

## GitHub Pages

Production файл `docs/` дотор байна. Шинэчлэхдээ:

```bash
npm run build
# Windows: docs хавтсыг dist-ээс дахин хуулна
```

Репозитори → **Settings** → **Pages** → **Source**: Deploy from a branch → `main` / `/docs`.

## Файлын бүтэц

```
mongolia-weekly-price-dashboard/
├── index.html
├── package.json
├── vite.config.js
├── public/favicon.svg
└── src/
    ├── main.js
    ├── app.js      # өгөгдөл, график, газрын зураг
    └── style.css
```

## Технологи

- [Vite](https://vite.dev/) + vanilla JavaScript
- [Chart.js](https://www.chartjs.org/)
- [Leaflet](https://leafletjs.com/) + CARTO dark tiles
- ҮСХ PxWeb API

## Зөвшөөрөл

Өгөгдөл — Үндэсний статистикийн хороо (1212.mn).  
Веб код — чөлөөтэй ашиглаж, засах боломжтой.
