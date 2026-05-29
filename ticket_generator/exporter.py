import os
from collections import Counter, defaultdict
from datetime import datetime
from typing import Iterable, List, Sequence

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .models import (
    BLOCK_PRACTICE,
    BLOCK_TEST,
    BLOCK_THEMATIC,
    CategoryGeneration,
    GenerationParams,
    Question,
    Ticket,
)


MONTHS_RU = {
    1: "января",
    2: "февраля",
    3: "марта",
    4: "апреля",
    5: "мая",
    6: "июня",
    7: "июля",
    8: "августа",
    9: "сентября",
    10: "октября",
    11: "ноября",
    12: "декабря",
}


HEADERS = [
    "Тема",
    "Название темы",
    "Вопрос",
    "Варианты ответов",
    "Правильный ответ",
    "Идентификатор",
]


def russian_date(date_value: datetime) -> str:
    return "%d %s %d" % (date_value.day, MONTHS_RU[date_value.month], date_value.year)


def category_filename(category: int, date_value: datetime) -> str:
    return "Билеты %d категория %s.xlsx" % (category, russian_date(date_value))


def stats_filename(date_value: datetime) -> str:
    return "Статистика генерации %s.xlsx" % russian_date(date_value)


def export_category_file(
    generation: CategoryGeneration, output_dir: str, date_value: datetime
) -> str:
    workbook = Workbook()
    default = workbook.active
    workbook.remove(default)

    for ticket in generation.tickets:
        worksheet = workbook.create_sheet("Билет %d" % ticket.number)
        _write_ticket_sheet(worksheet, ticket)

    identifiers = workbook.create_sheet("Идентификаторы")
    _write_identifiers_sheet(identifiers, generation.tickets)

    path = os.path.join(output_dir, category_filename(generation.category, date_value))
    workbook.save(path)
    return path


def export_stats_file(
    generations: Iterable[CategoryGeneration],
    output_dir: str,
    source_path: str,
    params: GenerationParams,
    date_value: datetime,
) -> str:
    generations = list(generations)
    workbook = Workbook()

    summary = workbook.active
    summary.title = "Сводка"
    _write_summary(summary, generations, source_path, params, date_value)

    main_overlap = workbook.create_sheet("Основные пересечения")
    _write_main_overlap(main_overlap, generations, params)

    theme_sheet = workbook.create_sheet("Темы")
    _write_theme_stats(theme_sheet, generations)

    repeated_sheet = workbook.create_sheet("Повторы тестовых")
    _write_test_repeats(repeated_sheet, generations)

    warnings_sheet = workbook.create_sheet("Предупреждения")
    _write_warnings(warnings_sheet, generations)

    path = os.path.join(output_dir, stats_filename(date_value))
    workbook.save(path)
    return path


def _write_summary(
    worksheet,
    generations: Sequence[CategoryGeneration],
    source_path: str,
    params: GenerationParams,
    date_value: datetime,
) -> None:
    worksheet.append(["Параметр", "Значение"])
    rows = [
        ("Исходный реестр", source_path),
        ("Дата генерации", date_value.strftime("%d.%m.%Y %H:%M")),
        ("Seed генерации", params.seed),
        ("Количество билетов", params.ticket_count),
        ("Основных билетов", params.main_ticket_count),
        ("Тестовых в билете", params.test_count),
        ("Тематических в билете", params.thematic_count),
        ("Практических в билете", params.practice_count),
        ("Максимум тестовых из одной темы в билете", params.max_questions_per_theme),
    ]
    for row in rows:
        worksheet.append(list(row))

    worksheet.append([])
    worksheet.append(
        [
            "Категория",
            "Основные билеты",
            "Тестовых в реестре",
            "Тестовых позиций во всех билетах",
            "Уникальных тестовых во всех билетах",
            "Покрытие базы",
            "Макс. использований вопроса",
            "Уникальных тестовых в основных",
            "Повторов в основных",
            "Макс. пересечение основных",
        ]
    )
    for generation in generations:
        test_source = _source_test_questions(generation)
        all_usage = _usage(generation.tickets, BLOCK_TEST)
        main_tickets = _main_tickets(generation)
        main_usage = _usage(main_tickets, BLOCK_TEST)
        max_main_overlap = _max_test_overlap(main_tickets)
        main_slots = len(main_tickets) * params.test_count
        worksheet.append(
            [
                generation.category,
                ", ".join(str(ticket.number) for ticket in main_tickets),
                len(test_source),
                params.ticket_count * params.test_count,
                len(all_usage),
                _percent(len(all_usage), len(test_source)),
                max(all_usage.values()) if all_usage else 0,
                len(main_usage),
                max(0, main_slots - len(main_usage)),
                max_main_overlap,
            ]
        )
    _style_table(worksheet)


def _write_main_overlap(
    worksheet, generations: Sequence[CategoryGeneration], params: GenerationParams
) -> None:
    worksheet.append(
        [
            "Категория",
            "Билет A",
            "Билет B",
            "Общих тестовых",
            "Пересечение от билета",
            "Общих тем",
            "Повторяющиеся ID тестовых",
        ]
    )
    for generation in generations:
        for left, right in _ticket_pairs(_main_tickets(generation)):
            common_ids = sorted(
                set(left.ids(BLOCK_TEST)) & set(right.ids(BLOCK_TEST)), key=int
            )
            common_topics = {
                question.topic
                for question in left.questions[BLOCK_TEST]
                if question.question_id in common_ids
            }
            worksheet.append(
                [
                    generation.category,
                    left.number,
                    right.number,
                    len(common_ids),
                    _percent(len(common_ids), params.test_count),
                    len(common_topics),
                    ", ".join(common_ids),
                ]
            )
    _style_table(worksheet)


def _write_theme_stats(worksheet, generations: Sequence[CategoryGeneration]) -> None:
    worksheet.append(
        [
            "Категория",
            "Тема",
            "Название темы",
            "Вопросов в реестре",
            "Уникально использовано во всех билетах",
            "Покрытие темы",
            "Использований во всех билетах",
            "Использований в основных",
            "Билетов с темой",
            "Макс. в одном билете",
        ]
    )
    for generation in generations:
        source_by_topic = defaultdict(list)
        for question in _source_test_questions(generation):
            source_by_topic[question.topic].append(question)
        all_tickets = generation.tickets
        main_tickets = _main_tickets(generation)
        for topic in sorted(source_by_topic, key=_topic_sort_key):
            source_questions = source_by_topic[topic]
            topic_name = source_questions[0].topic_name
            all_used_ids = {
                question.question_id
                for ticket in all_tickets
                for question in ticket.questions[BLOCK_TEST]
                if question.topic == topic
            }
            all_uses = sum(
                1
                for ticket in all_tickets
                for question in ticket.questions[BLOCK_TEST]
                if question.topic == topic
            )
            main_uses = sum(
                1
                for ticket in main_tickets
                for question in ticket.questions[BLOCK_TEST]
                if question.topic == topic
            )
            tickets_with_topic = sum(
                1
                for ticket in all_tickets
                if any(question.topic == topic for question in ticket.questions[BLOCK_TEST])
            )
            max_in_ticket = max(
                (
                    sum(1 for question in ticket.questions[BLOCK_TEST] if question.topic == topic)
                    for ticket in all_tickets
                ),
                default=0,
            )
            worksheet.append(
                [
                    generation.category,
                    topic,
                    topic_name,
                    len(source_questions),
                    len(all_used_ids),
                    _percent(len(all_used_ids), len(source_questions)),
                    all_uses,
                    main_uses,
                    tickets_with_topic,
                    max_in_ticket,
                ]
            )
    _style_table(worksheet)


def _write_test_repeats(
    worksheet, generations: Sequence[CategoryGeneration]
) -> None:
    worksheet.append(
        [
            "Категория",
            "Идентификатор",
            "Тема",
            "Название темы",
            "Использований во всех билетах",
            "Использований в основных",
            "В основных билетах",
        ]
    )
    for generation in generations:
        source_by_id = {question.question_id: question for question in _source_test_questions(generation)}
        all_usage = _usage(generation.tickets, BLOCK_TEST)
        main_tickets = _main_tickets(generation)
        main_usage = _usage(main_tickets, BLOCK_TEST)
        for question_id in sorted(source_by_id, key=int):
            if all_usage[question_id] <= 1 and main_usage[question_id] <= 1:
                continue
            question = source_by_id[question_id]
            main_numbers = [
                str(ticket.number)
                for ticket in main_tickets
                if question_id in ticket.ids(BLOCK_TEST)
            ]
            worksheet.append(
                [
                    generation.category,
                    question_id,
                    question.topic,
                    question.topic_name,
                    all_usage[question_id],
                    main_usage[question_id],
                    ", ".join(main_numbers),
                ]
            )
    _style_table(worksheet)


def _write_warnings(worksheet, generations: Sequence[CategoryGeneration]) -> None:
    worksheet.append(["Категория", "Предупреждение"])
    for generation in generations:
        for warning in generation.warnings:
            worksheet.append([generation.category, warning])
    _style_table(worksheet)


def _write_ticket_sheet(worksheet, ticket: Ticket) -> None:
    worksheet.append(HEADERS)
    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
    for question in ticket.all_questions():
        worksheet.append(_question_row(question))
    _style_table(worksheet)


def _question_row(question: Question) -> List[str]:
    topic = question.topic
    if question.block == BLOCK_THEMATIC:
        topic = "Тематический вопрос"
    elif question.block == BLOCK_PRACTICE:
        topic = "Практическая задача"
    return [
        topic,
        question.topic_name,
        question.question,
        question.answers,
        question.correct_answer,
        question.question_id,
    ]


def _write_identifiers_sheet(worksheet, tickets: List[Ticket]) -> None:
    rows_per_ticket = 6
    left_col = 1
    right_col = 3
    for index, ticket in enumerate(tickets):
        side_col = left_col if index < 5 else right_col
        block_index = index if index < 5 else index - 5
        row = block_index * rows_per_ticket + 1
        title = "Билет №%d" % ticket.number
        if ticket.is_main:
            title += " ★"
        worksheet.cell(row, side_col, title)
        worksheet.cell(row, side_col).font = Font(bold=True)
        worksheet.cell(row + 1, side_col, "Тестовые:")
        worksheet.cell(row + 1, side_col + 1, ", ".join(ticket.ids(BLOCK_TEST)))
        worksheet.cell(row + 2, side_col, "Тематические:")
        worksheet.cell(row + 2, side_col + 1, ", ".join(ticket.ids(BLOCK_THEMATIC)))
        worksheet.cell(row + 3, side_col, "Практические:")
        worksheet.cell(row + 3, side_col + 1, ", ".join(ticket.ids(BLOCK_PRACTICE)))
    _style_table(worksheet)
    worksheet.column_dimensions["A"].width = 18
    worksheet.column_dimensions["B"].width = 95
    worksheet.column_dimensions["C"].width = 18
    worksheet.column_dimensions["D"].width = 95


def _source_test_questions(generation: CategoryGeneration) -> List[Question]:
    return [question for question in generation.source_questions if question.block == BLOCK_TEST]


def _main_tickets(generation: CategoryGeneration) -> List[Ticket]:
    return [ticket for ticket in generation.tickets if ticket.is_main]


def _usage(tickets: Sequence[Ticket], block: str) -> Counter:
    usage = Counter()
    for ticket in tickets:
        for question in ticket.questions[block]:
            usage[question.question_id] += 1
    return usage


def _ticket_pairs(tickets: Sequence[Ticket]):
    for index, left in enumerate(tickets):
        for right in tickets[index + 1 :]:
            yield left, right


def _max_test_overlap(tickets: Sequence[Ticket]) -> int:
    max_overlap = 0
    for left, right in _ticket_pairs(tickets):
        overlap = len(set(left.ids(BLOCK_TEST)) & set(right.ids(BLOCK_TEST)))
        max_overlap = max(max_overlap, overlap)
    return max_overlap


def _percent(value: int, total: int) -> str:
    if not total:
        return "0%"
    return "%.1f%%" % (value / total * 100)


def _topic_sort_key(topic: str):
    return (0, int(topic)) if str(topic).isdigit() else (1, str(topic))


def _style_table(worksheet) -> None:
    if worksheet.max_row >= 1:
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
    for row in worksheet.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    widths = {
        1: 16,
        2: 18,
        3: 38,
        4: 22,
        5: 24,
        6: 18,
        7: 28,
        8: 22,
        9: 18,
        10: 18,
    }
    for col_idx in range(1, worksheet.max_column + 1):
        worksheet.column_dimensions[get_column_letter(col_idx)].width = widths.get(
            col_idx, 24
        )
    worksheet.freeze_panes = "A2"
