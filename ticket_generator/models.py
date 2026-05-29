from dataclasses import dataclass
from typing import Dict, List, Sequence


BLOCK_TEST = "test"
BLOCK_THEMATIC = "thematic"
BLOCK_PRACTICE = "practice"


BLOCK_LABELS = {
    BLOCK_TEST: "Тестовые",
    BLOCK_THEMATIC: "Тематические",
    BLOCK_PRACTICE: "Практические",
}


@dataclass(frozen=True)
class Question:
    category: int
    block: str
    topic: str
    topic_name: str
    question: str
    answers: str
    correct_answer: str
    question_id: str
    source_sheet: str
    source_row: int


@dataclass
class GenerationParams:
    ticket_count: int = 10
    main_ticket_count: int = 4
    test_count: int = 50
    thematic_count: int = 3
    practice_count: int = 2
    theme_fraction_limit: float = 0.5
    max_questions_per_theme: int = 4
    seed: int = 0


@dataclass
class Ticket:
    number: int
    questions: Dict[str, List[Question]]
    is_main: bool = False

    def all_questions(self) -> List[Question]:
        return (
            self.questions.get(BLOCK_TEST, [])
            + self.questions.get(BLOCK_THEMATIC, [])
            + self.questions.get(BLOCK_PRACTICE, [])
        )

    def ids(self, block: str) -> List[str]:
        return [q.question_id for q in self.questions.get(block, [])]


@dataclass
class CategoryGeneration:
    category: int
    tickets: List[Ticket]
    warnings: List[str]
    source_questions: List[Question]


@dataclass
class RegistryData:
    questions_by_category: Dict[int, List[Question]]
    warnings: List[str]


@dataclass
class ValidationResult:
    errors: List[str]
    warnings: List[str]
    counts: Dict[int, Dict[str, int]]

    @property
    def ok(self) -> bool:
        return not self.errors


def sort_questions(questions: Sequence[Question]) -> List[Question]:
    return sorted(questions, key=lambda q: int(q.question_id))
