# Assesment

Скилл для Claude, который проводит **продуктовый ассесмент PM/PO** по модели продуктовой зрелости
(опросник «Ассесмент PM V17 2026»).

## Как работает

1. Сотрудник подключается к скиллу.
2. Скилл просит перечислить все цели прошлого PI и уточняет их тип (продуктовая / регуляторка / тех. долг).
3. Скилл проводит интервью по опроснику в 4 блоках: **Гипотезы → ЦА → Метрики → Цели и стратегия**
   (блок «Рынок» пока не используется), по возможности привязывая вопросы к целям PI.
4. Баллы ставятся по правилам опросника и вкладки «Вводные данные»: накопление по уровням,
   пропорциональность при частичном выполнении, обязательные артефакты там, где их требует правило.
5. В конце скрипт заполняет копию Excel-шаблона:
   - листы срезов: баллы (желтые ячейки) и комментарии (голубые);
   - **«Радар»**: оценка данного PM и лепестковая диаграмма;
   - **«Зоны роста»**: зона, рекомендация, аргументация и общая рекомендация.

## Структура

```
product-assessment/
├── SKILL.md                      # инструкции скилла: клиентский путь, формат результатов
├── assets/assessment_template.xlsx
├── references/
│   ├── methodology.md            # правила из «Вводных данных» и «Идей к ассесменту»
│   ├── questionnaire.md          # вопросы, уровни, правила начисления с id
│   └── questionnaire.json
├── scripts/
│   ├── build_questionnaire.py    # шаблон Excel → questionnaire.json/.md
│   └── fill_assessment.py        # результаты JSON → заполненный Excel (+ отчет .md)
└── examples/sample_results.json
```

## Установка

**Claude.ai / Claude Desktop:** заархивируйте папку `product-assessment` в zip и загрузите архив
в *Settings → Capabilities → Skills*. Также должно быть включено выполнение кода (Code execution).

```bash
cd product-assessment/.. && zip -r product-assessment.zip product-assessment
```

**Claude Code:** скопируйте папку в `~/.claude/skills/` (личные скиллы) или в `.claude/skills/` проекта.

После этого достаточно написать Claude: «Хочу пройти продуктовый ассесмент».

## Проверка скрипта

```bash
pip install lxml openpyxl
```

```bash
python3 product-assessment/scripts/fill_assessment.py product-assessment/examples/sample_results.json -o sample.xlsx --report sample.md
```

## Если изменился шаблон

Положите новую версию в `product-assessment/assets/assessment_template.xlsx` и перегенерируйте опросник:

```bash
python3 product-assessment/scripts/build_questionnaire.py
```

Скрипт выведет суммы максимальных баллов по срезам. Сверьте их с максимумами из «Вводных данных» (25 / 16 / 24 / 12).
