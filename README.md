# DOCX Image Swap

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)
![Version](https://img.shields.io/badge/Version-1.0-orange)

[English](#english) | [Русский](#русский)

---

## English

Extract images from Word documents (.docx), replace them manually, and reinject back — without touching the text. Available as both **CLI** and **GUI** (ttkbootstrap).

### Features

- **Extract** all images from DOCX (inline + floating)
- **Manual swap** — replace image files in the output folder
- **Reinject** images back into their exact positions
- Auto JPEG optimization (configurable quality)
- DPI normalization (fixes stretched images)
- Aspect ratio preservation on reinsertion
- Configurable max dimensions (px and inches)
- Full document structure preservation (text, styles, formatting untouched)
- GUI mode with step-by-step pipeline visualization

### How It Works

```
 DOCX ──► Extract ──► images/ folder ──► Manual Swap ──► Reinject ──► Output DOCX
          (Step 1)                        (Step 2)        (Step 3)     (Step 4)
```

1. **Step 1** — Extract images from `source/*.docx` into `images/img1.jpg`, `img2.jpg`, ...
2. **Step 2** — Replace image files in `images/` folder (keep the same filenames)
3. **Step 3** — Reinject the new images back into the document
4. **Step 4** — Save the result to `output/`

### Installation

```bash
# Clone the repository
git clone https://github.com/al-nemirov/docx-image-swap.git
cd docx-image-swap

# Install dependencies
pip install -r requirements.txt

# (Optional) For GUI mode
pip install ttkbootstrap
```

### Usage

#### CLI mode

```bash
# Process the first .docx found in source/
python run.py

# Process a specific file
python run.py document.docx

# Process a file by full path
python run.py /path/to/my_document.docx
```

#### GUI mode

```bash
# Launch the graphical interface
python run.pyw
```

In GUI mode:
1. Click **Browse** to select a `.docx` file
2. Click **Run All Steps** to start the pipeline
3. When prompted, open the `images/` folder, replace image files, then click **Continue after swap**
4. Find the result in the `output/` folder

### Project Structure

```
docx-image-swap/
├── config.json                        — pipeline settings
├── run.py                             — CLI runner
├── run.pyw                            — GUI runner (ttkbootstrap)
├── requirements.txt                   — Python dependencies
├── modules/
│   ├── step_01_extract_images.py      — extract all images from DOCX
│   ├── step_02_manual_swap.py         — manual replacement pause
│   ├── step_03_insert_images.py       — reinject images into DOCX
│   └── step_04_save_result.py         — save output document
├── source/                            — input DOCX files (gitignored)
├── images/                            — extracted/replaced images (gitignored)
├── temp/                              — working directory (gitignored)
└── output/                            — result files (gitignored)
```

### Configuration

All settings are stored in `config.json`. Each step can be enabled/disabled independently.

| Field | Default | Description |
|-------|---------|-------------|
| `jpeg_quality` | `95` | JPEG compression quality (1-100) |
| `max_width` / `max_height` | `4096` | Max image dimensions in pixels |
| `max_width_inches` / `max_height_inches` | `5.5` / `8.0` | Max image size in the document (inches) |
| `center_images` | `true` | Center images horizontally on reinsertion |
| `preserve_aspect_ratio` | `true` | Maintain original proportions when resizing |
| `create_zip` | `false` | Create a ZIP archive of the output |
| `filename_template` | `{name}_img_swap_{timestamp}.docx` | Output filename pattern |

### Screenshots

<!-- Add screenshots of the GUI here -->
> Screenshots coming soon.

### Requirements

- Python 3.8+
- [python-docx](https://python-docx.readthedocs.io/) — Word document manipulation
- [Pillow](https://pillow.readthedocs.io/) — image processing
- [ttkbootstrap](https://ttkbootstrap.readthedocs.io/) — GUI only (optional)

### Contributing

Contributions are welcome! To contribute:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/my-feature`)
3. Commit your changes (`git commit -m "Add my feature"`)
4. Push to the branch (`git push origin feature/my-feature`)
5. Open a Pull Request

Please make sure your code follows the existing style and includes appropriate docstrings.

---

## Русский

Извлечение изображений из Word-документов (.docx), ручная замена и вставка обратно — без изменения текста. Доступно в режиме **CLI** и **GUI** (ttkbootstrap).

### Возможности

- **Извлечение** всех изображений из DOCX (встроенные + плавающие)
- **Ручная замена** — замените файлы в папке
- **Обратная вставка** — новые изображения возвращаются на свои места
- Авто-оптимизация JPEG (настраиваемое качество)
- Нормализация DPI (исправляет растянутые изображения)
- Сохранение пропорций при вставке
- Настраиваемые максимальные размеры (пкс и дюймы)
- Полное сохранение структуры документа (текст, стили, форматирование)
- GUI-режим с визуализацией пошагового процесса

### Как работает

```
 DOCX ──► Извлечь ──► images/ ──► Ручная замена ──► Вставить ──► Итоговый DOCX
          (Шаг 1)                  (Шаг 2)           (Шаг 3)     (Шаг 4)
```

1. **Шаг 1** — Извлечь изображения из `source/*.docx` в `images/img1.jpg`, `img2.jpg`, ...
2. **Шаг 2** — Заменить файлы в папке `images/` (сохранить имена файлов)
3. **Шаг 3** — Вставить новые изображения обратно в документ
4. **Шаг 4** — Сохранить результат в `output/`

### Установка

```bash
# Клонировать репозиторий
git clone https://github.com/al-nemirov/docx-image-swap.git
cd docx-image-swap

# Установить зависимости
pip install -r requirements.txt

# (Опционально) Для GUI-режима
pip install ttkbootstrap
```

### Использование

#### Режим CLI

```bash
# Обработать первый .docx в source/
python run.py

# Обработать конкретный файл
python run.py document.docx

# Обработать файл по полному пути
python run.py /path/to/my_document.docx
```

#### Режим GUI

```bash
# Запустить графический интерфейс
python run.pyw
```

В GUI-режиме:
1. Нажмите **Browse** для выбора `.docx` файла
2. Нажмите **Run All Steps** для запуска
3. Когда будет предложено, откройте папку `images/`, замените файлы изображений, нажмите **Continue after swap**
4. Результат будет в папке `output/`

### Структура проекта

```
docx-image-swap/
├── config.json                        — настройки пайплайна
├── run.py                             — CLI-запуск
├── run.pyw                            — GUI-запуск (ttkbootstrap)
├── requirements.txt                   — зависимости Python
├── modules/
│   ├── step_01_extract_images.py      — извлечение изображений из DOCX
│   ├── step_02_manual_swap.py         — пауза для ручной замены
│   ├── step_03_insert_images.py       — вставка изображений обратно
│   └── step_04_save_result.py         — сохранение результата
├── source/                            — входные DOCX (gitignored)
├── images/                            — извлечённые/заменённые (gitignored)
├── temp/                              — рабочая директория (gitignored)
└── output/                            — результат (gitignored)
```

### Настройка

Все параметры хранятся в `config.json`. Каждый шаг можно включить/отключить независимо.

| Поле | По умолч. | Описание |
|------|-----------|----------|
| `jpeg_quality` | `95` | Качество сжатия JPEG (1-100) |
| `max_width` / `max_height` | `4096` | Макс. размеры изображения (пкс) |
| `max_width_inches` / `max_height_inches` | `5.5` / `8.0` | Макс. размеры в документе (дюймы) |
| `center_images` | `true` | Центрировать при вставке |
| `preserve_aspect_ratio` | `true` | Сохранять пропорции при масштабировании |
| `create_zip` | `false` | Создать ZIP-архив результата |
| `filename_template` | `{name}_img_swap_{timestamp}.docx` | Шаблон имени выходного файла |

### Скриншоты

<!-- Добавьте скриншоты GUI здесь -->
> Скриншоты скоро будут добавлены.

### Требования

- Python 3.8+
- [python-docx](https://python-docx.readthedocs.io/) — работа с Word-документами
- [Pillow](https://pillow.readthedocs.io/) — обработка изображений
- [ttkbootstrap](https://ttkbootstrap.readthedocs.io/) — только для GUI (опционально)

### Участие в разработке

Мы приветствуем вклад в проект! Чтобы внести изменения:

1. Сделайте форк репозитория
2. Создайте ветку (`git checkout -b feature/my-feature`)
3. Зафиксируйте изменения (`git commit -m "Add my feature"`)
4. Отправьте ветку (`git push origin feature/my-feature`)
5. Откройте Pull Request

Пожалуйста, следуйте существующему стилю кода и добавляйте документацию к функциям.

---

## Author / Автор

**Alexander Nemirov** — [GitHub](https://github.com/al-nemirov)

## License / Лицензия

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

Проект распространяется по лицензии MIT. Подробности в файле [LICENSE](LICENSE).
