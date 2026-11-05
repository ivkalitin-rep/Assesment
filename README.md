# Assesment

Скиллы для Claude, которые проводят **продуктовый ассесмент PM/PO** мини-продуктов МЕТА по модели продуктовой
зрелости (опросник «Ассесмент PM V17 2026»).

В репозитории два варианта — выберите один:

| Папка | Вариант | Когда подходит |
|---|---|---|
| [`single-skill/`](single-skill) | **Единый скилл** `product-assessment`: интервью и оценка в одном скилле | самопроверка; ассесор сам ведёт интервью в своём чате; методика не секрет |
| [`two-skills/`](two-skills) | **Два раздельных скилла**: `product-assessment-interview` (сотрудникам) и `product-assessment-scoring` (ассесорам) | сотрудник проходит интервью сам и не должен видеть правила оценки |

## Как это работает

### Единый скилл

1. Знакомство: ФИО, команда, роль, мини-продукты, PI.
2. Цели PI.
3. Вопросы по одному: Гипотезы → ЦА → Метрики → Цели и стратегия (блок «Рынок» не используется).
4. Баллы по правилам опросника и методички с проверкой артефактов.
5. Excel: листы срезов, «Радар», «Зоны роста» + отчёт `.md`.

### Два скилла

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

В интервьюере нет правил начисления, баллов, описаний уровней и шаблона Excel: он ничего не оценивает.

## Структура

```
template/assessment_template.xlsx      # мастер-шаблон Excel (правьте только его)
tools/build_questionnaire.py           # шаблон → файлы вопросов и копии шаблона во всех скиллах

single-skill/product-assessment/       # единый скилл
├── SKILL.md
├── assets/assessment_template.xlsx    # копия мастер-шаблона
├── references/{methodology.md, questionnaire.md, questionnaire.json}
├── scripts/fill_assessment.py
└── examples/sample_results.json

two-skills/
├── product-assessment-interview/      # раздаётся сотрудникам
│   ├── SKILL.md
│   ├── references/{interview_rules.md, questions.md, questions.json}
│   ├── scripts/check_protocol.py
│   └── examples/sample_protocol.json
└── product-assessment-scoring/        # только для ассесоров
    ├── SKILL.md
    ├── assets/assessment_template.xlsx
    ├── references/{methodology.md, questionnaire.md, questionnaire.json}
    ├── scripts/fill_assessment.py
    └── examples/{sample_protocol.json, sample_results.json}
```

## Установка

**Claude.ai / Claude Desktop:** заархивируйте папку нужного скилла в zip и загрузите в *Settings → Capabilities → Skills*.
Нужно включённое выполнение кода (Code execution). Например, для единого скилла:

```bash
cd single-skill && zip -r product-assessment.zip product-assessment
```

Для двух скиллов:

```bash
cd two-skills && zip -r product-assessment-interview.zip product-assessment-interview && zip -r product-assessment-scoring.zip product-assessment-scoring
```

**Claude Code:** скопируйте папку скилла в `~/.claude/skills/`.

Вариант с двумя скиллами: сотрудникам раздаётся **только** `product-assessment-interview`, ассесорам — `product-assessment-scoring`.

## Ограничения варианта с двумя скиллами

- Протокол — файл у сотрудника, и до передачи ассесору его можно отредактировать. Интервьюер просит приложить
  ссылку на чат интервью, если среда позволяет ей поделиться.
- Описания артефактов в протоколе делает интервьюер. Ассесор может попросить оригиналы и сверить.
- Вопросы интервью сотруднику видны. Закрыты правила начисления и баллы.

## Общие правила и синхронизация

Правила интервью (один вопрос, ответ обязателен, мини-продукты, без подсказок, артефакты вместо ссылок, перерыв)
и методика оценки описаны в обоих вариантах. **Меняя правила в одном варианте, переносите их в другой:**

| Что меняется | Единый скилл | Два скилла |
|---|---|---|
| Ход интервью, вопросы, перерыв | `single-skill/product-assessment/SKILL.md` | `two-skills/product-assessment-interview/SKILL.md`, `references/interview_rules.md` |
| Как считать баллы, зоны роста | `single-skill/product-assessment/references/methodology.md` | `two-skills/product-assessment-scoring/references/methodology.md` |
| Заполнение Excel | `single-skill/product-assessment/scripts/fill_assessment.py` | `two-skills/product-assessment-scoring/scripts/fill_assessment.py` |

Вопросы и правила начисления из Excel синхронизирует генератор.

## Если изменился шаблон

Положите новую версию в `template/assessment_template.xlsx` и запустите:

```bash
python3 tools/build_questionnaire.py
```

Скрипт обновит копии шаблона и файлы вопросов во всех трёх скиллах и выведет суммы максимальных баллов по срезам.
Сверьте их с максимумами из «Вводных данных» (25 / 16 / 24 / 12).

## Проверка скриптов

```bash
pip install lxml openpyxl
```

```bash
python3 single-skill/product-assessment/scripts/fill_assessment.py single-skill/product-assessment/examples/sample_results.json -o sample.xlsx
```

```bash
python3 two-skills/product-assessment-interview/scripts/check_protocol.py two-skills/product-assessment-interview/examples/sample_protocol.json
```

```bash
python3 two-skills/product-assessment-scoring/scripts/fill_assessment.py two-skills/product-assessment-scoring/examples/sample_results.json -o sample.xlsx
```
