import os
from openpyxl import Workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def get_next_run_number():
    existing = [
        d for d in os.listdir(BASE_DIR)
        if os.path.isdir(os.path.join(BASE_DIR, d)) and d.startswith("batch run ")
    ]
    return len(existing) + 1


def gday_australia():
    run_number = get_next_run_number()
    folder = os.path.join(BASE_DIR, f"batch run {run_number}")
    os.makedirs(folder)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Batch Run {run_number}"
    ws.append(["Iteration", "Output"])

    for i in range(1, 101):
        message = f"G'day Australia! {i}"
        print(message)
        ws.append([i, message])

    wb.save(os.path.join(folder, "output.xlsx"))
    print(f"\nOutput saved to: batch run {run_number}/output.xlsx")


if __name__ == "__main__":
    gday_australia()
