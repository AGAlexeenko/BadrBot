import tkinter as tk


def display_list():
    # Пример списка
    items = ["Элемент 1", "Элемент 2", "Элемент 3", "Элемент 4"]

    # Очистить виджет
    listbox.delete(0, tk.END)

    # Добавить элементы в Listbox
    for item in items:
        listbox.insert(tk.END, item)


# Создаем основное окно
root = tk.Tk()
root.title("Вывод списка")

# Создаем виджет Listbox
listbox = tk.Listbox(root, height=10, width=30)
listbox.pack()

# Кнопка для отображения списка
button = tk.Button(root, text="Показать список", command=display_list)
button.pack()

# Запускаем главный цикл
root.mainloop()