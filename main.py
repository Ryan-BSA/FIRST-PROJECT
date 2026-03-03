import os
import tkinter as tk
from tkinter import messagebox
from openpyxl import Workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def get_next_run_number():
    existing = [
        d for d in os.listdir(BASE_DIR)
        if os.path.isdir(os.path.join(BASE_DIR, d)) and d.startswith("batch run ")
    ]
    return len(existing) + 1


def run(iterations, log_text):
    run_number = get_next_run_number()
    folder = os.path.join(BASE_DIR, f"batch run {run_number}")
    os.makedirs(folder)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Batch Run {run_number}"
    ws.append(["Iteration", "Output"])

    log_text.delete("1.0", tk.END)

    for i in range(1, iterations + 1):
        message = f"G'day Australia! {i}"
        ws.append([i, message])
        log_text.insert(tk.END, message + "\n")
        log_text.see(tk.END)
        log_text.update()

    filename = f"batch run {run_number}.xlsx"
    wb.save(os.path.join(folder, filename))
    log_text.insert(tk.END, f"\nOutput saved to: batch run {run_number}/{filename}\n")


def on_run():
    try:
        iterations = int(entry.get())
        if iterations < 1:
            raise ValueError
    except ValueError:
        messagebox.showerror("Invalid input", "Please enter a positive whole number.")
        return
    run(iterations, log_text)


app = tk.Tk()
app.title("G'day Australia")
app.resizable(False, False)

frame = tk.Frame(app, padx=10, pady=10)
frame.pack()

tk.Label(frame, text="Number of iterations:").grid(row=0, column=0, sticky="w")
entry = tk.Entry(frame, width=10)
entry.insert(0, "100")
entry.grid(row=0, column=1, padx=5)

tk.Button(frame, text="Run", command=on_run, width=10).grid(row=0, column=2)

log_text = tk.Text(frame, width=50, height=20, state="normal")
log_text.grid(row=1, column=0, columnspan=3, pady=(10, 0))

app.mainloop()
