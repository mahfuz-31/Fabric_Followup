import tkinter as tk
from tkinter import ttk
import pandas as pd
def fabric_data_window(root):
	new_window = tk.Toplevel(root)
	new_window.title("Color wise fabric status")
	new_window.geometry("1100x720")

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
			order_sheet_recL = tk.Label(window_frame, text="O. Sheet: " + row["Order Sheet Receive Date"], font=("Calibri", 9))
			cut_startL = tk.Label(window_frame, text="Cut S:" + row["Cut Plan Start Date"], font=("Calibri", 9))
			cut_endL = tk.Label(window_frame, text="Cut E: " + row["Cut Plan End Date"], font=("Calibri", 9))
			buyerL = tk.Label(window_frame, text=row["Buyer"], font=("Calibri", 11, 'bold'), fg="darkgreen")

			buyerL.grid(row=row_idx, column=0, sticky="w", padx=2, pady=2)
			orderL.grid(row=row_idx, column=1, sticky="w", padx=2, pady=2)
			styleL.grid(row=row_idx, column=2, sticky="w", padx=2, pady=2)
			order_sheet_recL.grid(row=row_idx, column=3, sticky="w", padx=2, pady=2)
			cut_startL.grid(row=row_idx, column=4, sticky="w", padx=2, pady=2)
			cut_endL.grid(row=row_idx, column=5, sticky="w", padx=2, pady=2)
			# Add only bottom line
			row_idx += 1
			separator = tk.Frame(window_frame, height=2, bd=0, relief=tk.FLAT, bg="black")
			separator.grid(row=row_idx, column=0, columnspan=9, sticky="ew", pady=2)
			# add the below table header
			row_idx += 1
			headers = ["F. Color", "G. Color", "UoF", "F. Type", "GSM", "Gray Book.","Gray Status", "F/F Book.", "F/F Status"]
			for col_idx, header in enumerate(headers):
				headerL = tk.Label(window_frame, text=header, font=("Calibri", 11, 'normal', 'underline'), fg="darkgreen")
				headerL.grid(row=row_idx, column=col_idx, sticky="w", padx=2, pady=2)
			prev_order = order
			row_idx += 1

		# Now add the row data
		color_wise_data = [
			row["F. Color"],
			row["G. Color"],
			row["UoF"],
			row["F. Type"],
			row["GSM"],
		]
		col_i = 0
		for col_idx, value in enumerate(color_wise_data):
			dataL = tk.Label(window_frame, text=value, font=("Calibri", 11))
			dataL.grid(row=row_idx, column=col_idx, sticky="w", padx=2, pady=2)
			col_i = col_idx

		# gray fabric data
		col_i += 1
		gray_fabric_data = [
			row["G/F Order With S.Note Qty"],
			row["Net Grey Receive Qty"],
			row["G/F Rcv Balance Qty"],
		]
		gbookL = tk.Label(window_frame, text=gray_fabric_data[0], font=("Calibri", 11))
		gbookL.grid(row=row_idx, column=col_i, sticky="w", padx=2, pady=2)
		col_i += 1
		gray_progress = ttk.Progressbar(window_frame, orient='horizontal', length=120, mode='determinate')
		gray_progress['maximum'] = gray_fabric_data[0] if gray_fabric_data[0] > 0 else 1  # avoid zero division
		gray_progress['value'] = gray_fabric_data[1]
		gray_progress.grid(row=row_idx, column=col_i, padx=2, pady=2)
		
		# F/F fabric data
		col_i += 1
		ff_fabric_data = [
			row["F/F Order with S.Note Qty"],
			row["F/F Delv Qty"],
			row["F/F Delv. Balance Qty"],
		]
		ffbookL = tk.Label(window_frame, text=ff_fabric_data[0], font=("Calibri", 11))
		ffbookL.grid(row=row_idx, column=col_i, sticky="w", padx=2, pady=2)
		col_i += 1
		ff_progress = ttk.Progressbar(window_frame, orient='horizontal', length=120, mode='determinate')
		ff_progress['maximum'] = ff_fabric_data[0] if ff_fabric_data[0] > 0 else 1	# avoid zero division
		ff_progress['value'] = ff_fabric_data[1]
		ff_progress.grid(row=row_idx, column=col_i, padx=2, pady=2)
		
		row_idx += 1

	# --- Enable mouse wheel scrolling ---
	def _on_mouse_wheel(event):
			canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

	canvas.bind_all("<MouseWheel>", _on_mouse_wheel)   # Windows
	canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))  # Linux up
	canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))   # Linux down



        