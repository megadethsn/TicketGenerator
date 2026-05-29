import os
import sys
import traceback
from datetime import datetime
from tkinter import (
    BooleanVar,
    Button,
    Checkbutton,
    Entry,
    Frame,
    IntVar,
    Label,
    LabelFrame,
    Listbox,
    Tk,
    W,
    filedialog,
    messagebox,
)

from ticket_generator.exporter import (
    category_filename,
    export_category_file,
    export_stats_file,
    stats_filename,
)
from ticket_generator.generator import GenerationError, generate_category
from ticket_generator.models import (
    BLOCK_PRACTICE,
    BLOCK_TEST,
    BLOCK_THEMATIC,
    GenerationParams,
)
from ticket_generator.registry_reader import read_registry


class TicketGeneratorApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("Генератор экзаменационных билетов")
        self.root.geometry("900x650")

        self.registry_path = None
        self.output_dir = None
        self.ticket_count = IntVar(value=10)
        self.test_count = IntVar(value=50)
        self.thematic_count = IntVar(value=3)
        self.practice_count = IntVar(value=2)
        self.category_vars = {category: BooleanVar(value=True) for category in range(1, 9)}
        self.log = None

        self._build_ui()

    def _build_ui(self) -> None:
        files = LabelFrame(self.root, text="Файлы", padx=8, pady=8)
        files.pack(fill="x", padx=10, pady=8)

        self.registry_path = Entry(files, width=85)
        self.output_dir = Entry(files, width=85)

        Label(files, text="Реестр Excel").grid(row=0, column=0, sticky=W, pady=3)
        self.registry_path.grid(row=0, column=1, sticky=W, pady=3)
        Button(files, text="Выбрать", command=self._choose_registry).grid(row=0, column=2, padx=5)

        Label(files, text="Папка результата").grid(row=1, column=0, sticky=W, pady=3)
        self.output_dir.grid(row=1, column=1, sticky=W, pady=3)
        Button(files, text="Выбрать", command=self._choose_output_dir).grid(row=1, column=2, padx=5)

        options = Frame(self.root)
        options.pack(fill="x", padx=10, pady=4)

        categories = LabelFrame(options, text="Категории", padx=8, pady=8)
        categories.pack(side="left", fill="y", padx=(0, 8))
        for category in range(1, 9):
            Checkbutton(
                categories,
                text="Категория %d" % category,
                variable=self.category_vars[category],
            ).grid(row=(category - 1) // 2, column=(category - 1) % 2, sticky=W, padx=4)
        Button(categories, text="Все", command=self._select_all_categories).grid(row=4, column=0, sticky=W, pady=6)
        Button(categories, text="Снять", command=self._clear_categories).grid(row=4, column=1, sticky=W, pady=6)

        params = LabelFrame(options, text="Параметры", padx=8, pady=8)
        params.pack(side="left", fill="both", expand=True)
        self._add_int_entry(params, 0, "Количество билетов", self.ticket_count)
        self._add_int_entry(params, 1, "Тестовых вопросов", self.test_count)
        self._add_int_entry(params, 2, "Тематических вопросов", self.thematic_count)
        self._add_int_entry(params, 3, "Практических задач", self.practice_count)
        Label(params, text="Основных билетов: 4").grid(row=4, column=0, sticky=W, pady=4)

        actions = Frame(self.root)
        actions.pack(fill="x", padx=10, pady=8)
        Button(actions, text="Проверить реестр", command=self._validate).pack(side="left", padx=(0, 8))
        Button(actions, text="Сформировать билеты", command=self._generate).pack(side="left")

        log_frame = LabelFrame(self.root, text="Сообщения", padx=8, pady=8)
        log_frame.pack(fill="both", expand=True, padx=10, pady=8)
        self.log = Listbox(log_frame)
        self.log.pack(fill="both", expand=True)

    def _add_int_entry(self, parent, row: int, label: str, var: IntVar) -> None:
        Label(parent, text=label).grid(row=row, column=0, sticky=W, pady=4)
        Entry(parent, textvariable=var, width=8).grid(row=row, column=1, sticky=W, pady=4)

    def _choose_registry(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите реестр вопросов",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.registry_path.delete(0, "end")
            self.registry_path.insert(0, path)
            if not self.output_dir.get():
                self.output_dir.insert(0, os.path.dirname(path))

    def _choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Выберите папку сохранения")
        if path:
            self.output_dir.delete(0, "end")
            self.output_dir.insert(0, path)

    def _select_all_categories(self) -> None:
        for var in self.category_vars.values():
            var.set(True)

    def _clear_categories(self) -> None:
        for var in self.category_vars.values():
            var.set(False)

    def _selected_categories(self):
        return [category for category, var in self.category_vars.items() if var.get()]

    def _params(self) -> GenerationParams:
        params = GenerationParams(
            ticket_count=int(self.ticket_count.get()),
            main_ticket_count=4,
            test_count=int(self.test_count.get()),
            thematic_count=int(self.thematic_count.get()),
            practice_count=int(self.practice_count.get()),
        )
        if params.ticket_count < 4:
            raise ValueError("Количество билетов должно быть не меньше 4.")
        if params.test_count < 1:
            raise ValueError("Количество тестовых вопросов должно быть больше 0.")
        if params.thematic_count < 0 or params.practice_count < 0:
            raise ValueError("Количество вопросов не может быть отрицательным.")
        return params

    def _validate(self):
        self._clear_log()
        try:
            categories = self._selected_categories()
            if not categories:
                raise ValueError("Выберите хотя бы одну категорию.")
            registry_path = self.registry_path.get().strip()
            if not registry_path:
                raise ValueError("Выберите Excel-реестр.")
            params = self._params()

            registry, result = read_registry(registry_path, categories)
            errors = list(result.errors)
            for category in categories:
                counts = result.counts.get(category, {})
                if counts.get(BLOCK_TEST, 0) < params.test_count:
                    errors.append(
                        "Категория %d: тестовых вопросов %d, требуется %d."
                        % (category, counts.get(BLOCK_TEST, 0), params.test_count)
                    )
                if counts.get(BLOCK_THEMATIC, 0) < params.thematic_count:
                    errors.append(
                        "Категория %d: тематических вопросов %d, требуется %d."
                        % (category, counts.get(BLOCK_THEMATIC, 0), params.thematic_count)
                    )
                if counts.get(BLOCK_PRACTICE, 0) < params.practice_count:
                    errors.append(
                        "Категория %d: практических задач %d, требуется %d."
                        % (category, counts.get(BLOCK_PRACTICE, 0), params.practice_count)
                    )

            for category in categories:
                counts = result.counts.get(category, {})
                self._log(
                    "Категория %d: тестовые %d, тематические %d, практические %d."
                    % (
                        category,
                        counts.get(BLOCK_TEST, 0),
                        counts.get(BLOCK_THEMATIC, 0),
                        counts.get(BLOCK_PRACTICE, 0),
                    )
                )

            if errors:
                self._log("")
                self._log("Ошибки:")
                for error in errors:
                    self._log(error)
                messagebox.showerror("Проверка не пройдена", "Найдены ошибки. Генерация запрещена.")
                return False

            self._log("")
            self._log("Проверка пройдена. Можно формировать билеты.")
            messagebox.showinfo("Проверка", "Проверка пройдена.")
            return True
        except Exception as exc:
            self._log("Ошибка: %s" % exc)
            messagebox.showerror("Ошибка", str(exc))
            return False

    def _generate(self):
        self._clear_log()
        try:
            categories = self._selected_categories()
            if not categories:
                raise ValueError("Выберите хотя бы одну категорию.")
            registry_path = self.registry_path.get().strip()
            output_dir = self.output_dir.get().strip()
            if not registry_path:
                raise ValueError("Выберите Excel-реестр.")
            if not output_dir:
                raise ValueError("Выберите папку сохранения.")
            if not os.path.isdir(output_dir):
                raise ValueError("Папка сохранения не найдена.")

            params = self._params()
            date_value = datetime.now()
            planned_files = [
                os.path.join(output_dir, category_filename(category, date_value))
                for category in categories
            ]
            planned_files.append(os.path.join(output_dir, stats_filename(date_value)))
            existing = [path for path in planned_files if os.path.exists(path)]
            if existing:
                answer = messagebox.askyesno(
                    "Файлы уже существуют",
                    "Некоторые файлы уже существуют. Перезаписать?\n\n%s"
                    % "\n".join(existing),
                )
                if not answer:
                    self._log("Генерация отменена пользователем.")
                    return

            registry, result = read_registry(registry_path, categories)
            if result.errors:
                self._log("Ошибки реестра:")
                for error in result.errors:
                    self._log(error)
                messagebox.showerror("Ошибка", "Найдены ошибки реестра. Генерация запрещена.")
                return

            generations = []
            for category in categories:
                self._log("Формирование категории %d..." % category)
                generation = generate_category(
                    category, registry.questions_by_category.get(category, []), params
                )
                path = export_category_file(generation, output_dir, date_value)
                self._log("Создан файл: %s" % path)
                for warning in generation.warnings:
                    self._log("Предупреждение: %s" % warning)
                generations.append(generation)

            stats_path = export_stats_file(
                generations, output_dir, registry_path, params, date_value
            )
            self._log("Создан файл статистики: %s" % stats_path)
            messagebox.showinfo("Готово", "Билеты сформированы.")
        except (ValueError, GenerationError) as exc:
            self._log("Ошибка: %s" % exc)
            messagebox.showerror("Ошибка", str(exc))
        except Exception as exc:
            self._log("Непредвиденная ошибка: %s" % exc)
            self._log(traceback.format_exc())
            messagebox.showerror("Ошибка", str(exc))

    def _clear_log(self) -> None:
        self.log.delete(0, "end")

    def _log(self, message: str) -> None:
        self.log.insert("end", message)
        self.log.see("end")
        self.root.update_idletasks()


def main() -> int:
    root = Tk()
    TicketGeneratorApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
