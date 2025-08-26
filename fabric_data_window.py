import tkinter as tk
import pandas as pd
def fabric_data_window(root):
	new_window = tk.Toplevel(root)
	new_window.title("New Window")
	new_window.geometry("1000x780")
	# new_window.configure(bg="lightyellow")
	data = pd.read_excel("color_wise_output.xlsx")
	orders = data["Order"].unique()
	prev_order = orders[0]
	row_idx = 0
	for index, row in data.iterrows():
		order = row["Order"]
		if order != prev_order or index == 0:
			orderL = tk.Label(new_window, text=row["Order"], font=("Arial", 12, "bold"), fg="blue")
			styleL = tk.Label(new_window, text=row["Style"], font=("Arial", 12, "bold"), fg="blue")
			orderL.grid(row=row_idx, column=0)
			styleL.grid(row=row_idx, column=1)
			prev_order = order
			row_idx += 1
		FcolorL = tk.Label(new_window, text=row["F. Color"], )
		FcolorL.grid(row=row_idx, column=0)
		print(row["F. Color"], row_idx)

		row_idx += 1


        