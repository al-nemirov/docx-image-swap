# DOCX Image Swap

[Русский](#русский) | [English](#english)

---

## English

Extract images from Word documents (.docx), replace them manually, and reinject back — without touching the text.

### Features

- **Extract** all images from DOCX (inline + floating)
- **Manual swap** — replace image files in the folder
- **Reinject** images back into exact positions
- Auto JPEG optimization (configurable quality)
- DPI normalization (fixes stretched images)
- Aspect ratio preservation on reinsertion
- Configurable max dimensions
- Full document structure preservation (text, styles, formatting untouched)

### How It Works

1. **Step 1:** Extract images from `source/*.docx` → saved as `images/img1.jpg`, `img2.jpg`...
2. **Step 2:** Replace image files in `images/` folder (keep same filenames)
3. **Step 3:** Reinject new images back into the document
4. **Step 4:** Save result to `output/`

### Project Structure

```
docx-image-swap/
├── config.json                        — settings
├── modules/
│   ├── step_01_extract_images.py      — extract all images
│   ├── step_02_manual_swap.py         — manual replacement step
│   ├── step_03_insert_images.py       — reinject images
│   └── step_04_save_result.py         — save output
├── source/                            — input DOCX (gitignored)
├── images/                            — extracted/replaced images (gitignored)
└── output/                            — result (gitignored)
```

### Requirements

- Python 3.8+
- `python-docx`
- `Pillow`

### Configuration

Edit `config.json`:

| Field | Default | Description |
|-------|---------|-------------|
| `jpeg_quality` | 95 | JPEG compression quality |
| `max_width` / `max_height` | 4096 | Max image dimensions (px) |
| `max_width_inches` / `max_height_inches` | 5.5 / 8.0 | Max size in document |
| `center_images` | true | Center images on reinsertion |
| `preserve_aspect_ratio` | true | Maintain proportions |

---

## Русский

Извлечение изображений из Word-документов (.docx), ручная замена и вставка обратно — без изменения текста.

### Возможности

- **Извлечение** всех изображений из DOCX (встроенные + плавающие)
- **Ручная замена** — замените файлы в папке
- **Обратная вставка** — новые изображения возвращаются на свои места
- Авто-оптимизация JPEG (настраиваемое качество)
- Нормализация DPI (исправляет растянутые изображения)
- Сохранение пропорций при вставке
- Настраиваемые максимальные размеры
- Полное сохранение структуры документа (текст, стили, форматирование)

### Как работает

1. **Шаг 1:** Извлечь изображения из `source/*.docx` → сохраняются как `images/img1.jpg`, `img2.jpg`...
2. **Шаг 2:** Заменить файлы в папке `images/` (имена сохранить)
3. **Шаг 3:** Вставить новые изображения обратно в документ
4. **Шаг 4:** Сохранить результат в `output/`

### Структура

```
docx-image-swap/
├── config.json                        — настройки
├── modules/
│   ├── step_01_extract_images.py      — извлечение изображений
│   ├── step_02_manual_swap.py         — ручная замена
│   ├── step_03_insert_images.py       — вставка обратно
│   └── step_04_save_result.py         — сохранение результата
├── source/                            — входные DOCX (gitignored)
├── images/                            — извлечённые/заменённые (gitignored)
└── output/                            — результат (gitignored)
```

### Требования

- Python 3.8+
- `python-docx`
- `Pillow`

### Настройка

Отредактируйте `config.json`:

| Поле | По умолч. | Описание |
|------|-----------|----------|
| `jpeg_quality` | 95 | Качество сжатия JPEG |
| `max_width` / `max_height` | 4096 | Макс. размеры изображения (пкс) |
| `max_width_inches` / `max_height_inches` | 5.5 / 8.0 | Макс. размеры в документе |
| `center_images` | true | Центрировать при вставке |
| `preserve_aspect_ratio` | true | Сохранять пропорции |

## Author / Автор

Alexander Nemirov
