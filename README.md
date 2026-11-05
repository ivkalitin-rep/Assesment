# Assesment

Скиллы для Claude, которые проводят **продуктовый ассесмент PM/PO** мини-продуктов МЕТА по модели продуктовой
зрелости (опросник «Ассесмент PM V17 2026»).

> **Ветки.** В `main` лежит вариант из одного скилла `product-assessment`: интервью и оценка вместе.
> В этой ветке — вариант из **двух скиллов**: вопросы и правила оценки разделены, чтобы сотрудник не видел методику оценки.

## Как это работает

```
Сотрудник                                   Ассесор
─────────                                   ───────
product-assessment-interview                product-assessment-scoring
  1. Знакомство, мини-продукты, цели PI       5. Принимает протокол и артефакты
  2. Вопросы по одному: Гипотезы → ЦА →       6. Ставит баллы по правилам опросника
     Метрики → Цели и стратегия               7. Зоны роста и общая рекомендация
  3. Собирает ответы и артефакты              8. Заполняет Excel: «Радар», «Зоны роста»
  4. Выдаёт протокол (json + md + zip)  ──►   9. Обратную связь сотруднику даёт ассесор
```

**Интервьюер** (`product-assessment-interview`) — для сотрудников. Содержит только вопросы и правила ведения
интервью. Правил начисления, баллов, описаний уровней и шаблона Excel в нём нет. Он ничего не оценивает.

**Оценщик** (`product-assessment-scoring`) — только для ассесоров. Содержит методику, правила начисления,
Excel-шаблон и скрипт заполнения.

## Структура

```
product-assessment-interview/          # раздаётся сотрудникам
├── SKILL.md                           # ход интервью, перерыв, формат протокола
├── references/
│   ├── interview_rules.md             # как вести интервью (без правил оценки)
│   ├── questions.md                   # вопросы с пометками [артефакт] и [если «нет» — почему?]
│   └── questions.json
├── scripts/check_protocol.py          # проверка полноты протокола, md-версия и zip для ассесора
└── examples/sample_protocol.json

product-assessment-scoring/            # только для ассесоров
├── SKILL.md                           # как оценивать протокол и заполнять Excel
├── assets/assessment_template.xlsx
├── references/
│   ├── methodology.md                 # «Вводные данные», «Идеи к ассесменту», как считать балл
│   ├── questionnaire.md               # вопросы, уровни, правила начисления с id
│   └── questionnaire.json
├── scripts/fill_assessment.py         # результаты JSON → заполненный Excel (+ отчёт .md)
└── examples/{sample_protocol.json, sample_results.json}

tools/build_questionnaire.py           # шаблон Excel → файлы вопросов для обоих скиллов
```

## Установка

**Claude.ai / Claude Desktop:** заархивируйте папку нужного скилла в zip и загрузите в *Settings → Capabilities → Skills*.
Нужно включённое выполнение кода (Code execution).

```bash
zip -r product-assessment-interview.zip product-assessment-interview
```

```bash
zip -r product-assessment-scoring.zip product-assessment-scoring
```

**Claude Code:** скопируйте папку скилла в `~/.claude/skills/`.

Кому что раздавать:
- сотрудникам — **только** `product-assessment-interview`;
- ассесорам — `product-assessment-scoring` (и при желании интервьюер, чтобы видеть, как идёт интервью).

## Ограничения

- Протокол — это файл у сотрудника, и до передачи ассесору его можно отредактировать. Чтобы сверять спорные места,
  интервьюер просит сотрудника приложить ссылку на чат интервью, если среда позволяет ей поделиться.
- Описания артефактов в протоколе делает интервьюер. Ассесор может попросить оригиналы и сверить.
- Вопросы интервью сотруднику видны: файлы интервьюера открыты так же, как любой скилл. Закрыты правила и баллы.

## Проверка скриптов

```bash
pip install lxml openpyxl
```

```bash
python3 product-assessment-interview/scripts/check_protocol.py product-assessment-interview/examples/sample_protocol.json
```

```bash
python3 product-assessment-scoring/scripts/fill_assessment.py product-assessment-scoring/examples/sample_results.json -o sample.xlsx --report sample.md
```

## Если изменился шаблон

Положите новую версию в `product-assessment-scoring/assets/assessment_template.xlsx` и перегенерируйте вопросы
для обоих скиллов:

```bash
python3 tools/build_questionnaire.py
```

Скрипт выведет суммы максимальных баллов по срезам. Сверьте их с максимумами из «Вводных данных» (25 / 16 / 24 / 12).
