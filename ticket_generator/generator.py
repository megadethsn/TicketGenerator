import itertools
import math
import random
from collections import Counter, defaultdict
from typing import Dict, List, Sequence, Set

from .models import (
    BLOCK_PRACTICE,
    BLOCK_TEST,
    BLOCK_THEMATIC,
    CategoryGeneration,
    GenerationParams,
    Question,
    Ticket,
    sort_questions,
)


class GenerationError(Exception):
    pass


def generate_category(
    category: int, questions: Sequence[Question], params: GenerationParams
) -> CategoryGeneration:
    pools = {
        BLOCK_TEST: sort_questions([q for q in questions if q.block == BLOCK_TEST]),
        BLOCK_THEMATIC: sort_questions([q for q in questions if q.block == BLOCK_THEMATIC]),
        BLOCK_PRACTICE: sort_questions([q for q in questions if q.block == BLOCK_PRACTICE]),
    }
    _ensure_pool(category, "тестовых вопросов", pools[BLOCK_TEST], params.test_count)
    _ensure_pool(
        category, "тематических вопросов", pools[BLOCK_THEMATIC], params.thematic_count
    )
    _ensure_pool(
        category, "практических задач", pools[BLOCK_PRACTICE], params.practice_count
    )
    if params.ticket_count < params.main_ticket_count:
        raise GenerationError(
            "Категория %d: количество билетов меньше 4 основных." % category
        )

    rng = random.Random(params.seed + category * 1009)
    tickets: List[Ticket] = []
    usage = {BLOCK_TEST: Counter(), BLOCK_THEMATIC: Counter(), BLOCK_PRACTICE: Counter()}
    combo_usage = {BLOCK_THEMATIC: Counter(), BLOCK_PRACTICE: Counter()}
    co_usage: Dict[str, Counter] = defaultdict(Counter)
    theme_counts_total = Counter(q.topic for q in pools[BLOCK_TEST])

    for number in range(1, params.ticket_count + 1):
        test_questions = _select_test_questions(
            pools[BLOCK_TEST],
            params.test_count,
            usage[BLOCK_TEST],
            co_usage,
            theme_counts_total,
            params.theme_fraction_limit,
            params.max_questions_per_theme,
            rng,
        )
        _remember_selection(test_questions, usage[BLOCK_TEST], co_usage)

        thematic_questions = _select_balanced_combo(
            pools[BLOCK_THEMATIC],
            params.thematic_count,
            usage[BLOCK_THEMATIC],
            combo_usage[BLOCK_THEMATIC],
            rng,
        )
        _remember_simple(thematic_questions, usage[BLOCK_THEMATIC], combo_usage[BLOCK_THEMATIC])

        practice_questions = _select_balanced_combo(
            pools[BLOCK_PRACTICE],
            params.practice_count,
            usage[BLOCK_PRACTICE],
            combo_usage[BLOCK_PRACTICE],
            rng,
        )
        _remember_simple(practice_questions, usage[BLOCK_PRACTICE], combo_usage[BLOCK_PRACTICE])

        tickets.append(
            Ticket(
                number=number,
                questions={
                    BLOCK_TEST: sort_questions(test_questions),
                    BLOCK_THEMATIC: sort_questions(thematic_questions),
                    BLOCK_PRACTICE: sort_questions(practice_questions),
                },
            )
        )

    for ticket in _choose_main_tickets(tickets, params.main_ticket_count):
        ticket.is_main = True

    warnings = _build_warnings(category, pools, params)
    return CategoryGeneration(
        category=category,
        tickets=tickets,
        warnings=warnings,
        source_questions=list(questions),
    )


def _ensure_pool(category: int, label: str, pool: Sequence[Question], required: int) -> None:
    if len(pool) < required:
        raise GenerationError(
            "Категория %d: недостаточно %s: доступно %d, требуется %d."
            % (category, label, len(pool), required)
        )


def _select_test_questions(
    pool: Sequence[Question],
    required: int,
    usage: Counter,
    co_usage: Dict[str, Counter],
    theme_counts_total: Counter,
    fraction_limit: float,
    max_questions_per_theme: int,
    rng: random.Random,
) -> List[Question]:
    selected: List[Question] = []
    selected_ids: Set[str] = set()
    theme_in_ticket = Counter()
    attempts = 0

    while len(selected) < required:
        attempts += 1
        if attempts > required * 5:
            # If strict theme limits made a rare configuration impossible, relax
            # only the current ticket. The validation still guarantees enough rows.
            fraction_limit = min(1.0, fraction_limit + 0.1)
            attempts = 0

        candidates = [
            q
            for q in pool
            if q.question_id not in selected_ids
            and theme_in_ticket[q.topic]
            < _theme_limit(q.topic, theme_counts_total, fraction_limit, max_questions_per_theme)
        ]
        if not candidates:
            candidates = [q for q in pool if q.question_id not in selected_ids]
        if not candidates:
            raise GenerationError("Не удалось подобрать тестовые вопросы без дублей внутри билета.")

        scored = []
        for question in candidates:
            pair_penalty = sum(co_usage[question.question_id][other.question_id] for other in selected)
            score = (
                usage[question.question_id] * 1000
                + theme_in_ticket[question.topic] * 80
                + pair_penalty * 25
                + rng.random()
            )
            scored.append((score, question))
        scored.sort(key=lambda item: item[0])
        chosen = scored[0][1]
        selected.append(chosen)
        selected_ids.add(chosen.question_id)
        theme_in_ticket[chosen.topic] += 1

    return selected


def _theme_limit(
    topic: str, theme_counts_total: Counter, fraction_limit: float, max_questions_per_theme: int
) -> int:
    total = theme_counts_total[topic]
    proportional_limit = max(1, int(math.ceil(total * fraction_limit)))
    if max_questions_per_theme <= 0:
        return proportional_limit
    return min(proportional_limit, max_questions_per_theme)


def _select_balanced_combo(
    pool: Sequence[Question],
    required: int,
    usage: Counter,
    combo_usage: Counter,
    rng: random.Random,
) -> List[Question]:
    if len(pool) == required:
        return list(pool)

    combinations = list(itertools.combinations(pool, required))
    scored = []
    for combo in combinations:
        combo_key = tuple(q.question_id for q in combo)
        score = (
            combo_usage[combo_key] * 1000
            + sum(usage[q.question_id] for q in combo) * 100
            + _internal_topic_repeats(combo) * 10
            + rng.random()
        )
        scored.append((score, combo))
    scored.sort(key=lambda item: item[0])
    return list(scored[0][1])


def _internal_topic_repeats(combo: Sequence[Question]) -> int:
    counts = Counter(q.topic for q in combo)
    return sum(value - 1 for value in counts.values() if value > 1)


def _remember_selection(
    questions: Sequence[Question], usage: Counter, co_usage: Dict[str, Counter]
) -> None:
    for question in questions:
        usage[question.question_id] += 1
    for left, right in itertools.combinations(questions, 2):
        co_usage[left.question_id][right.question_id] += 1
        co_usage[right.question_id][left.question_id] += 1


def _remember_simple(questions: Sequence[Question], usage: Counter, combo_usage: Counter) -> None:
    for question in questions:
        usage[question.question_id] += 1
    combo_usage[tuple(q.question_id for q in sort_questions(questions))] += 1


def _choose_main_tickets(tickets: Sequence[Ticket], count: int) -> List[Ticket]:
    best_combo = None
    best_score = None
    for combo in itertools.combinations(tickets, count):
        pair_scores = []
        test_overlaps = []
        unique_test_ids = set()
        test_slots = 0
        for ticket in combo:
            ticket_ids = ticket.ids(BLOCK_TEST)
            test_slots += len(ticket_ids)
            unique_test_ids.update(ticket_ids)
        for left, right in itertools.combinations(combo, 2):
            pair_scores.append(_ticket_similarity(left, right))
            test_overlaps.append(
                len(set(left.ids(BLOCK_TEST)) & set(right.ids(BLOCK_TEST)))
            )
        score = (
            test_slots - len(unique_test_ids),
            max(test_overlaps),
            sum(test_overlaps),
            max(pair_scores),
            sum(pair_scores),
        )
        if best_score is None or score < best_score:
            best_score = score
            best_combo = combo
    return list(best_combo or tickets[:count])


def _ticket_similarity(left: Ticket, right: Ticket) -> int:
    left_test = set(left.ids(BLOCK_TEST))
    right_test = set(right.ids(BLOCK_TEST))
    left_thematic = set(left.ids(BLOCK_THEMATIC))
    right_thematic = set(right.ids(BLOCK_THEMATIC))
    left_practice = set(left.ids(BLOCK_PRACTICE))
    right_practice = set(right.ids(BLOCK_PRACTICE))
    left_topics = {q.topic for q in left.questions[BLOCK_TEST]}
    right_topics = {q.topic for q in right.questions[BLOCK_TEST]}

    return (
        len(left_test & right_test) * 100
        + len(left_thematic & right_thematic) * 35
        + len(left_practice & right_practice) * 5
        + len(left_topics & right_topics)
    )


def _build_warnings(
    category: int, pools: Dict[str, Sequence[Question]], params: GenerationParams
) -> List[str]:
    warnings: List[str] = []
    if len(pools[BLOCK_PRACTICE]) == params.practice_count:
        warnings.append(
            "Категория %d: практические задачи будут повторяться во всех билетах, "
            "так как доступно ровно %d из %d."
            % (category, params.practice_count, params.practice_count)
        )
    if len(pools[BLOCK_THEMATIC]) < params.ticket_count:
        warnings.append(
            "Категория %d: тематические вопросы будут повторяться, так как доступно %d."
            % (category, len(pools[BLOCK_THEMATIC]))
        )
    return warnings
