# Лабораторная работа 2. Линейное программирование

Вариант 16. Требуется найти минимум и максимум целевой функции
`Z = -3x1 + 6x2` при заданных линейных ограничениях, решить задачу
графическим методом и методом симплекс-таблиц.

## Состав работы

| № | Тема | Ноутбук | Colab |
|---|---|---|---|
| 01 | Графический метод | [`01_graphical_method.ipynb`](notebooks/01_graphical_method.ipynb) | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/megusto0/mp-2/blob/main/notebooks/01_graphical_method.ipynb) |
| 02 | Симплекс-метод для максимума | [`02_simplex_max.ipynb`](notebooks/02_simplex_max.ipynb) | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/megusto0/mp-2/blob/main/notebooks/02_simplex_max.ipynb) |
| 03 | Симплекс-метод для минимума | [`03_simplex_min.ipynb`](notebooks/03_simplex_min.ipynb) | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/megusto0/mp-2/blob/main/notebooks/03_simplex_min.ipynb) |
| 04 | Сводная проверка результатов | [`04_summary.ipynb`](notebooks/04_summary.ipynb) | [![Open in Colab](https://colab.research.google.com/assets/colab-badge.svg)](https://colab.research.google.com/github/megusto0/mp-2/blob/main/notebooks/04_summary.ipynb) |

Ноутбуки являются самостоятельными: в них нет `git clone`, поэтому каждый файл можно открыть напрямую в Colab.

## Результат

Допустимая область имеет три вершины:

| Точка | x1 | x2 | Z |
|---|---:|---:|---:|
| A1 | 4/3 | 8/3 | 12 |
| A2 | 12/7 | 16/7 | 60/7 |
| A3 | 2 | 3 | 12 |

Итог:

- минимум: `x* = (12/7; 16/7)`, `Zmin = 60/7`;
- максимум: `Zmax = 12` на отрезке от `(4/3; 8/3)` до `(2; 3)`;
- симплекс-метод для максимума находит вершину `(2; 3)` и показывает альтернативный оптимум по нулевой оценке небазисной переменной.

## Структура проекта

- [`src/lp2`](src/lp2) - постановка задачи, графический метод, двухфазный симплекс-метод и построение графика;
- [`scripts/generate_artifacts.py`](scripts/generate_artifacts.py) - генерация таблиц и рисунка;
- [`scripts/generate_notebooks.py`](scripts/generate_notebooks.py) - генерация самостоятельных Colab-ноутбуков;
- [`scripts/build_report.py`](scripts/build_report.py) - генерация локального DOCX-отчета с формулами-изображениями;
- [`tables`](tables) - CSV/JSON с вычисленными результатами;
- [`figures`](figures) - рисунок допустимой области.

## Запуск локально

```powershell
pip install -r requirements.txt
python scripts\generate_artifacts.py
python scripts\generate_notebooks.py
python scripts\build_report.py
```

DOCX-отчет `report_lab2_v16_defense.docx` создается локально и не добавляется в git.
