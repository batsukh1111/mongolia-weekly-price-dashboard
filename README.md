# 7 хоногийн үнийн дашбоард (ҮСХ 1212.mn)

Хүнсний гол нэр болон бензин түлшний **7 хоногийн дундаж үнэ**-ийг Улаанбаатар болон 21 аймгаар харуулах интерактив HTML дашбоард.

## Онцлог

- **Ерөнхий** — KPI, хамгийн их өөрчлөлт, УБ үнийн түвшин, нийтлэг бүтээгдэхүүнээр аймаг vs УБ
- **Улаанбаатар** — 30+ бүтээгдэхүүн, чиг хандлага, 7 хоногийн % өөрчлөлт, хайлттай хүснэгт
- **Аймгууд** — 11 бүтээгдэхүүн × 21 аймаг, зэрэглэл, өөрчлөлт, дулаан зураг (heatmap)
- **Харьцуулалт** — УБ vs аймгийн дундаж / мин / макс (нийтлэг бүтээгдэхүүнүүд)
- **Чиг хандлага** — олон бүтээгдэхүүний индекс (эхлэл=100) болон нэрлэсэн үнэ
- **Амьд өгөгдөл** — хуудас нээгдэх / «Шинэчлэх» бүрт 1212.mn API-аас шууд татна (7 хоног бүр эх сурвалж шинэчлэгдэхэд дашбоард автоматаар дагана)

## Өгөгдлийн эх

| Бүс | Хүснэгт | API |
|-----|---------|-----|
| **Улаанбаатар** (бүтээгдэхүүн илүү олон) | `DT_NSO_0600_001V4` | [table-view](https://www.1212.mn/mn/statcate/table-view/Economy,%20environment/Consumer%20Price%20Index/DT_NSO_0600_001V4.px) |
| **Аймгууд** (11 бүтээгдэхүүн) | `DT_NSO_0300_010V5` | [table-view](https://www.1212.mn/mn/statcate/table-view/Economy,%20environment/Consumer%20Price%20Index/DT_NSO_0300_010V5.px) |

- API: [https://data.1212.mn/](https://data.1212.mn/) (PxWeb JSON-stat2)
- Нэгж: төгрөг (₮)
- Давтамж: 7 хоног

> Аймгуудын хүснэгтэд бүтээгдэхүүний нэр төрөл цөөн, нийслэлийнхэд илүү олон. «Харьцуулалт» хэсэгт зөвхөн хоёр эхэд нийтлэг байгаа 10 бүтээгдэхүүнийг харьцуулна.

## Локаль ажиллуулах

Цэвэр HTML тул framework шаардлагагүй. CORS-ийн улмаас `file://` биш, жижиг HTTP серверээр нээх нь найдвартай:

```bash
cd mongolia-weekly-price-dashboard

# Python
python -m http.server 8080

# эсвэл Node
npx serve .
```

Дараа нь: http://localhost:8080

## GitHub Pages дээр байрлуулах

### 1. Репозитори үүсгэх

1. [github.com/new](https://github.com/new) дээр шинэ **public** репозитори үүсгэнэ  
   (жнь: `mongolia-weekly-price-dashboard`)

### 2. Код оруулах

```bash
cd mongolia-weekly-price-dashboard
git init
git add index.html README.md
git commit -m "7 хоногийн үнийн дашбоард — 1212.mn амьд өгөгдөл"
git branch -M main
git remote add origin https://github.com/ТАНЫ_НЭР/mongolia-weekly-price-dashboard.git
git push -u origin main
```

### 3. Pages идэвхжүүлэх

1. Репозитори → **Settings** → **Pages**
2. **Source**: Deploy from a branch
3. **Branch**: `main` / `/ (root)` → Save
4. Хэдэн минутын дараа:

   `https://ТАНЫ_НЭР.github.io/mongolia-weekly-price-dashboard/`

Сервер тал дээр дата хадгалах шаардлагагүй — браузер шууд ҮСХ API руу хандана.

## Файлын бүтэц

```
mongolia-weekly-price-dashboard/
├── index.html   # UI + стиль + логик (нэг файл)
└── README.md
```

## Технологи

- HTML + CSS + JavaScript (framework байхгүй)
- [Chart.js](https://www.chartjs.org/) — график
- ҮСХ PxWeb API (CORS зөвшөөрсөн)

## Зөвшөөрөл

Өгөгдөл — Үндэсний статистикийн хороо (1212.mn).  
Веб код — чөлөөтэй ашиглаж, засах боломжтой.
