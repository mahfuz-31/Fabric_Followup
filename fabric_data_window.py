import tkinter as tk
import pandas as pd
def fabric_data_window(root):
	new_window = tk.Toplevel(root)
	new_window.title("New Window")
	new_window.geometry("1000x780")

	# --- Create Canvas + Scrollbar ---
	canvas = tk.Canvas(new_window)
	canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

	scrollbar = tk.Scrollbar(new_window, orient=tk.VERTICAL, command=canvas.yview)
	scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

	canvas.configure(yscrollcommand=scrollbar.set)

	# --- Frame inside Canvas ---
	window_frame = tk.Frame(canvas)
	canvas.create_window((0, 0), window=window_frame, anchor="nw")

	# Update scroll region whenever size changes
	def on_configure(event):
			canvas.configure(scrollregion=canvas.bbox("all"))

	window_frame.bind("<Configure>", on_configure)

	# --- Your Data ---
	data = pd.read_excel("color_wise_output.xlsx")
	orders = data["Order"].unique()
	prev_order = orders[0]
	row_idx = 0

	for index, row in data.iterrows():
		order = row["Order"]
		if order != prev_order or index == 0:
			orderL = tk.Label(window_frame, text=row["Order"], font=("Calibri", 12, "bold"), fg="blue")
			styleL = tk.Label(window_frame, text=row["Style"], font=("Calibri", 12, "bold"), fg="blue")
			order_sheet_recL = tk.Label(window_frame, text=row["Order Sheet Receive Date"], font=("Calibri", 11, 'bold'))
			cut_startL = tk.Label(window_frame, text=row["Cut Plan Start Date"], font=("Calibri", 11, 'bold'))
			cut_endL = tk.Label(window_frame, text=row["Cut Plan End Date"], font=("Calibri", 11, 'bold'))

			orderL.grid(row=row_idx, column=0, sticky="w", padx=5, pady=2)
			styleL.grid(row=row_idx, column=1, sticky="w", padx=5, pady=2)
			order_sheet_recL.grid(row=row_idx, column=2, sticky="w", padx=5, pady=2)
			cut_startL.grid(row=row_idx, column=3, sticky="w", padx=5, pady=2)
			cut_endL.grid(row=row_idx, column=4, sticky="w", padx=5, pady=2)
			# Add only bottom line
			row_idx += 1
			separator = tk.Frame(window_frame, height=2, bd=0, relief=tk.FLAT, bg="black")
			separator.grid(row=row_idx, column=0, columnspan=5, sticky="ew", pady=2)
			# add the below table header
			row_idx += 1
			headers = ["F. Color", "G. Color", "UoF", "F. Type", "GSM",]
			for col_idx, header in enumerate(headers):
				headerL = tk.Label(window_frame, text=header, font=("Calibri", 11, 'normal', 'underline'), fg="darkgreen")
				headerL.grid(row=row_idx, column=col_idx, sticky="w", padx=5, pady=2)
			prev_order = order
			row_idx += 1

		FcolorL = tk.Label(window_frame, text=row["F. Color"], font=("Calibri", 11, 'bold'))
		FcolorL.grid(row=row_idx, column=0, sticky="w", padx=5, pady=2)
		row_idx += 1

	# --- Enable mouse wheel scrolling ---
	def _on_mouse_wheel(event):
			canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

	canvas.bind_all("<MouseWheel>", _on_mouse_wheel)   # Windows
	canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))  # Linux up
	canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))   # Linux down



        