# created on 1 Mar 2022
from tkinter import *
from tkinter import filedialog, messagebox, simpledialog,ttk
import pandas as pd
import os, glob, gc, warnings
import gspread
from google.oauth2.service_account import Credentials

# connect to GSheet API
scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive']

credentials = Credentials.from_service_account_file('keys_drive.json',scopes=scopes)
gc = gspread.authorize(credentials)

sh = gc.open('testAPI')
worksheet = sh.sheet1

root = Tk()
root.geometry('800x600+1800+100')
root.title('My Withholding Tax')
root.resizable(True, True)
warnings.simplefilter('ignore')

class App_Inquiry:

	def __init__(self, master):
		self.frame = Frame(master)
		
		self.dev = Label(self.frame, text='DEMO Version: Created by Natcha Phonkamhaeng  ', fg='grey')
		self.journal_btn = Button(self.frame, text='  Import Cash Journal  ', command=self.browse_cash)
		self.inq_btn = Button(self.frame, text='  Import Etax Inquiry  ', command=self.browse_inq)
		self.merge_btn = Button(self.frame, text = 'Merge Journal & Inquiry', state=DISABLED, command=self.merge_data)
		self.clear_btn = Button(self.frame, text=' CLEAR ALL DATA', command=self.clear, fg='red')
		
		self.listbox_journal = Listbox(self.frame, selectmode=MULTIPLE, bd=1, height=25, width=30)
		self.listbox_inq = Listbox(self.frame, selectmode=MULTIPLE, bd=1, height=25, width=30)

		self.var_cash_btn = False
		self.var_inq_btn = False

	def browse_cash(self):			
		filetypes = [('Excel files', '.xlsx')]
		self.cash_path = filedialog.askopenfilenames(title='Choose files', filetypes=filetypes)

		self.df_cash = pd.DataFrame()

		for i in self.cash_path:
			self.cash_dir, self.cash_basename = os.path.split(i)
			self.listbox_journal.insert(END, self.cash_basename)
			# ETL by using Pandas
			for f in glob.glob(i):
				try:
					csh_chq = pd.read_excel(f, sheet_name='Report WHT Cash & CHQ')
					tfr = pd.read_excel(f, sheet_name='WHT TRF (HSBC&BAY)')
					e_payment = pd.read_excel(f, sheet_name='WHT TRF (e-payment)')
					df_concat = pd.concat([csh_chq, tfr, e_payment], ignore_index=True)
					self.df_cash = pd.concat([self.df_cash, df_concat], ignore_index=True)
				except ValueError as e:
					print(f'ValueError: file is not cash journal: {e}')
					messagebox.showerror('Error', f'File is not Cash Journal\n{e}')
					if 'ok':
						self.listbox_journal.delete(0,END)			
		try:
			# dropna on columns WHT
			self.df_cash.dropna(subset=['Unnamed: 10'], inplace=True)			
			# delete unnesscesary columns
			self.df_cash = self.df_cash.drop(self.df_cash.columns[12:], axis=1)
			# make row 1 to be header
			self.df_cash.columns = self.df_cash.iloc[0]
			self.df_cash = self.df_cash[1:]
			# delete unwanted row
			filt = (self.df_cash['Receipt Date'] == 'Receipt Date')
			self.df_cash = self.df_cash.drop(index=self.df_cash[filt].index)
			self.df_cash['Receipt Date'] = pd.to_datetime(self.df_cash['Receipt Date'], format='%Y-%m-%d')
			self.df_cash['Receipt\nApply AMT'] = self.df_cash['Receipt\nApply AMT'].astype(float)
			# df.to_excel('test.xlsx', index=False)
		except KeyError as e:
			print(f'KeyError: user did not select any file: {e}')
		except ValueError as e:
			print(f'ValueError: file is not cash journal: {e}')
			messagebox.showerror('Error', 'File is not Cash Journal')
		except Exception as e:
			print(f'Exception: error: {e}')

		# if listbox cash journal is not emtry (have data), browse button will be disable.
		if self.listbox_journal.get(0, END) != ():
			self.journal_btn['state'] = 'disabled'
			self.var_cash_btn = True
			self.switch()
											
	def browse_inq(self):
		filetypes = [('Excel files', '.xlsx')]
		
		self.inq_path = filedialog.askopenfilenames(title='Choose files', filetypes=filetypes)

		self.df_inq = pd.DataFrame()

		for i in self.inq_path:
			self.inq_dir, self.inq_basename = os.path.split(i)
			self.listbox_inq.insert(END, self.inq_basename)
			for f in glob.glob(i):
				read_inq = pd.read_excel(f, engine='openpyxl')
				self.df_inq = pd.concat([self.df_inq, read_inq], ignore_index=True)
		try:
			pd.set_option('display.float_format', '{:.0f}'.format)
			self.df_inq = self.df_inq[['docId', 'taxIDBR']]			
			self.df_inq['taxIDBR'] = self.df_inq['taxIDBR'].apply(lambda x: '{0:0>13}'.format(x))
			self.df_inq['taxIDBR'] = self.df_inq['taxIDBR'].apply(lambda x: x.split('.')[0])
			self.df_inq['taxIDBR'] = self.df_inq['taxIDBR'].astype(str)
			self.df_inq.drop_duplicates(subset=['docId'], inplace=True)
		# df_inq.to_excel("test_inq_all.xlsx", index=False)
		except KeyError as e:
			print(f'KeyError: user did not select file: {e}' )
			messagebox.showerror('Error', f'File is not Etax Inquiry\n{e}')
			if 'ok':
				self.listbox_inq.delete(0, END)		
		except Exception as e:
			print(f'Exception error: {e}')

		# if inq box cash journal is not emtry (have data), browse button will be disable.
		if self.listbox_inq.get(0, END) != ():
			self.inq_btn['state'] = 'disabled'
			self.var_inq_btn = True
			self.switch()
						
	def switch(self):
		if self.var_cash_btn and self.var_inq_btn:
			self.merge_btn['state'] = 'normal'
			self.var_cash_btn = False
			self.var_inq_btn = False
				
	def merge_data(self):			
		self.df_result = pd.merge(self.df_cash, self.df_inq, how='left', left_on='Receipt No', right_on='docId')
		self.df_result.drop(columns='docId', inplace=True)			
		self.df_result.rename(columns={
			'Bank\nAccount' : 'Bank',
			'Receipt\nCust Name' : 'Receipt Name',
			'Total Receipt\nAmount	' : 'Receipt Amount',
			'Receipt\nApply AMT': 'WHT Amount',
			}, inplace=True)

		self.df_result['Receipt Date'] = pd.to_datetime(self.df_result['Receipt Date'], format='%Y-%m-%d')
		self.df_result['WHT Amount'] = self.df_result['WHT Amount'].astype(float)
		
		self.frame.pack_forget()
		app_wht.draw()

	def draw(self):
		self.frame.pack(fill='both', expand=True)
		# grid_row/columnconfigure --> give ability to resize widget with position the same position button
		#(first arg is highest number of row and column)
		self.frame.grid_rowconfigure(3, weight=1)
		self.frame.grid_columnconfigure(2, weight=1)

		self.journal_btn.grid(row=0, column=0,pady=30, padx=60, sticky=W)
		self.listbox_journal.grid(row=1, column=0, padx=50, sticky=NW)
		self.inq_btn.grid(row=0, column=1, pady=30, padx=45, sticky=W)
		self.listbox_inq.grid(row=1, column=1,padx=40, sticky=NW)
		self.clear_btn.grid(row=0, column=2, pady=30,padx=35, sticky=W)
		self.merge_btn.grid(row=1, column=2, padx=35, sticky=W)
		self.dev.grid(row=3, column=0, pady=50, padx=30, sticky=SW)

	def clear(self):
		# clear all listbox item
		self.listbox_journal.delete(0, END)
		self.listbox_inq.delete(0, END)
		# grab all df and delete then set to emtry dataframe, delete all garbage collect
		try:
			del [[self.df_cash, self.df_inq]]
			gc.collect()
			self.df_cash = pd.DataFrame()
			self.df_inq = pd.DataFrame()
		except AttributeError as e:
			print(f'dataframe is not exist : {e}')
		
		self.journal_btn['state'] = 'normal'
		self.inq_btn['state'] = 'normal'
		self.merge_btn['state'] = 'disabled'

	# def drop_df_result(self):
	# 	del self.df_result
	# 	gc.collect()
		

class App_WHT:

	def __init__(self, master):
		self.frame = Frame(master)

		self.menubar = Menu(self.frame)
		self.filemenu = Menu(self.menubar, tearoff=0)
		self.filemenu.add_command(label='Download Excel', command=self.download)
		self.filemenu.add_command(label='Back to Main Menu', command=self.back )
		self.filemenu.add_separator()
		self.filemenu.add_command(label='Exit', command=root.quit)
		self.menubar.add_cascade(label='File', menu=self.filemenu)

		self.database_menu = Menu(self.menubar, tearoff=0)
		self.database_menu.add_command(label='Download from database', command=self.download_database)
		self.database_menu.add_command(label='Check With Cash Journal', command=self.check)
		self.database_menu.add_command(label='Pending from Cash Journal', command=self.pending)
		self.menubar.add_cascade(label='Query from Database', menu=self.database_menu)
		
		self.y_scrollbar = Scrollbar(self.frame)
		self.tree_view = ttk.Treeview(self.frame, height=20, selectmode='extended', yscrollcommand=self.y_scrollbar.set)
		
		self.search_label = Label(self.frame, text='Search BL')
		self.search_entry = Entry(self.frame)
		self.search_btn = Button(self.frame, text='  Search BL ', command=self.search)
		
		self.tax_label = Label(self.frame, text='Search Tax ID')
		self.tax_entry = Entry(self.frame)
		self.tax_btn = Button(self.frame, text=' Search Tax ', command=self.search_tax)

		self.cal_btn = Button(self.frame, text=' Cal WHT  ', command=self.cal_wht)
		self.cal_label = Label(self.frame, text='')

		self.clear_btn = Button(self.frame, text=' Clear All ', command=self.clear, anchor=N, fg='red')

		self.save_btn = Button(self.frame, text=' SAVE ',  bg='blue', fg='white', command=self.save,)

		self.remark_entry = Entry(self.frame)
		self.remark_label = Label(self.frame, text='Remark')
		self.remark_btn = Button(self.frame, text=' Remark ', command=self.remark)		

		self.adjust_label = Label(self.frame, text='Adjust')
		self.adjust_entry = Entry(self.frame)

	def download(self):
		filetypes = [('excel', '*.xlsx')]

		wht_data = filedialog.asksaveasfilename(filetypes=filetypes, defaultextension=filetypes, initialfile='wht_data')
		app_inquiry.df_result.to_excel(wht_data, index=False)

	def ttk_style(self, fbg):
		treeview_style = ttk.Style()
		treeview_style.theme_use('default')
		treeview_style.configure('Treeview',
			background='white',
			foreground='black',
			rowheight=25,
			fieldbackground=fbg
			)

		treeview_style.map('Treeview',
			background=[('selected', '#2623db')]
			)
		
		self.y_scrollbar.pack(side=RIGHT, fill=Y)
		
	def draw(self):
		self.frame.pack(fill='both', expand=True)
		root.config(menu=self.menubar)
		self.clear_treeview()
		# add style to Treeview first, if not strip line color won't work in tkinter
		self.ttk_style('white')
		
		# select column name that will show in Treeview
		filt = ['Receipt Date', 'Receipt Name', 'Receipt No', 'B/L No', 'WHT Amount', 'taxIDBR']
		self.tree_view['column'] = filt
		self.tree_view['show'] = 'headings'

		for col in self.tree_view['column']:
			self.tree_view.heading(col, text=col)

		self.df_rows = app_inquiry.df_result.filter(items=filt).to_numpy().tolist()		

		self.tree_view.tag_configure('oddrow', background='white')
		self.tree_view.tag_configure('evenrow', background='#d1edeb')

		count_row=0
		for row in self.df_rows:
			if count_row %2 == 0:
				self.tree_view.insert('', 'end', values=row, tags=('evenrow',))
			else:
				self.tree_view.insert('', 'end', values=row, tags=('oddrow',))
			count_row += 1

		self.y_scrollbar.config(command=self.tree_view.yview)
		self.tree_view.pack(fill=X)

		self.search_label.pack(padx=10, pady=10, side=LEFT, anchor=N)
		self.search_entry.pack(side=LEFT,pady=10, anchor=N)
		self.search_btn.pack(side=LEFT,pady=10, padx=8, anchor=N)

		self.tax_label.pack(padx=10, pady=10, side=LEFT, anchor=N)
		self.tax_entry.pack(side=LEFT, pady=10, anchor=N)
		self.tax_btn.pack(side=LEFT, pady=10, padx=8, anchor=N)				

		self.clear_btn.pack(padx=10, pady=10, side=LEFT, anchor=N)

		self.cal_label.pack(pady=10)
		self.save_btn.pack(padx=10, pady=10, side=RIGHT, anchor=N)
		self.remark_btn.pack(padx=10, pady=10, side=RIGHT, anchor=N)
		self.cal_btn.pack(padx=10, pady=10, side=RIGHT, anchor=N)

	def search_tax(self):
		try:
			query = int(self.tax_entry.get())			
			my_list = []
		
			for child in self.tree_view.get_children():
				if query == self.tree_view.item(child)['values'][5]:
					my_list.append(self.tree_view.item(child)['values'])

			self.clear_treeview()
			self.ttk_style('silver')

			filt = ['Receipt Date', 'Receipt Name', 'Receipt No', 'B/L No', 'WHT Amount', 'taxIDBR']
			self.tree_view['column'] = filt
			self.tree_view['show'] = 'headings'

			for col in self.tree_view['column']:
				self.tree_view.heading(col, text=col)

			self.tree_view.tag_configure('oddrow', background='white')
			self.tree_view.tag_configure('evenrow', background='#d1edeb')	
					
			count_row=0
			for row in my_list:
				if count_row %2 == 0:
					self.tree_view.insert('', 'end', values=row, tags=('evenrow',))
				else:
					self.tree_view.insert('', 'end', values=row, tags=('oddrow',))
				count_row += 1

			self.y_scrollbar.config(command=self.tree_view.yview)

		except ValueError as e:
			print(f'ValueError: search on emtry entry will return str: {e}')
						
	def search(self):
		query = self.search_entry.get().upper()
		my_list = []

		for child in self.tree_view.get_children():
			if query in self.tree_view.item(child)['values'][3]:				
				my_list.append(self.tree_view.item(child)['values'])

		self.clear_treeview()

		self.ttk_style('silver')

		filt = ['Receipt Date', 'Receipt Name', 'Receipt No', 'B/L No', 'WHT Amount', 'taxIDBR']
		self.tree_view['column'] = filt
		self.tree_view['show'] = 'headings'

		for col in self.tree_view['column']:
			self.tree_view.heading(col, text=col)

		self.tree_view.tag_configure('oddrow', background='white')
		self.tree_view.tag_configure('evenrow', background='#d1edeb')	
				
		count_row=0
		for row in my_list:
			if count_row %2 == 0:
				self.tree_view.insert('', 'end', values=row, tags=('evenrow',))
			else:
				self.tree_view.insert('', 'end', values=row, tags=('oddrow',))
			count_row += 1

		self.y_scrollbar.config(command=self.tree_view.yview)

	def cal_wht(self):
		wht_list = []

		for item in self.tree_view.selection():
			item_text = self.tree_view.item(item, 'values')[4]
			wht_list.append(item_text)

		wht_list = list(map(float, wht_list))
		sum_wht = sum(wht_list)
		
		self.cal_label.config(text=f'Total WHT: {sum_wht:,.2f}')

	def clear(self):
		self.draw()
		self.search_entry.delete(0, END)
		self.tax_entry.delete(0, END)
		self.cal_label.config(text='')
		self.cal_label.pack(pady=10)
		
			
	def clear_treeview(self):
		self.tree_view.delete(*self.tree_view.get_children())

	def back(self):
		self.frame.pack_forget()		
					
		emptyMenu = Menu(root)
		root.config(menu=emptyMenu)
		main()

	def save(self):
		wht_list = []

		for item in self.tree_view.selection():
			item_text = self.tree_view.item(item, 'values')[4]
			wht_list.append(item_text)

		wht_list = list(map(float, wht_list))
		sum_wht = sum(wht_list)

		confirm = messagebox.askyesno(title='Save to Database', message=f'Total WHT: {sum_wht:,.2f}\n\n Are you sure to save to Database?')
		if confirm:
			try:
				self.record()
				messagebox.showinfo(title='Save to Database', message='Save Successful')
				self.clear()
			except Exception as e:
				print(e)
				messagebox.showerror(title='Error', message=e)

	def remark(self):
		# first let user input all remark
		remark = simpledialog.askstring(title='Remark Reason', prompt='Input your remark')
		if remark != None:
			adjust = simpledialog.askfloat(title='Adjust', prompt='Amount to adjust\n\nIf no adjust, type 0.00')
			if adjust != None:
				confirm = messagebox.askyesno(title='Save to Database', message=f'Are you sure to save to Database?\n\nRemark : {remark}\nadjust : {adjust}')
				if confirm:
					messagebox.showinfo(title='Save to Database', message='Save Successful')					
					
		# make pandas dataframe for those remark
		list_remark = []
		list_remark.append(remark)

		list_adj = []
		list_adj.append(adjust)

		combine = list(zip(list_remark, list_adj))
		df_rm = pd.DataFrame(combine)
		df_rm.columns = df_rm.iloc[0]
		df_rm.drop(index=0, inplace=True)	
		# make pandas dataframe for selection in treeview
		data_list = []
		for item in self.tree_view.selection():
			item_text = self.tree_view.item(item, 'values')		
			data_list.append(item_text)

		df = pd.DataFrame(data_list)
		df.columns = df.iloc[0]
		df.drop(index=0, inplace=True)
		# combine two dataframe together using pd.concat
		df_concat = pd.concat([df, df_rm], axis=1)
		# send dataframe to API
		try:
			worksheet.append_rows([df_concat.columns.values.tolist()] + df_concat.values.tolist(), value_input_option="USER_ENTERED")		
		except Exception as e:
			print(e)
			messagebox.showerror(title='API Error', message='Please select one receipt for remark')
							
	def record(self):
		data_list = []
		for item in self.tree_view.selection():
			item_text = self.tree_view.item(item, 'values')		
			data_list.append(item_text)

		df = pd.DataFrame(data_list)
		df.columns = df.iloc[0]
		df.drop(index=0, inplace=True)
		
		worksheet.append_rows([df.columns.values.tolist()] + df.values.tolist() , value_input_option="USER_ENTERED")

	def gdata(self):
		get_gsheet = worksheet.get_all_values()
		self.df = pd.DataFrame(get_gsheet)
		self.df.columns = self.df.iloc[0]
		self.df.drop(index=0, inplace=True)
		self.df['WHT'] = self.df['WHT'].astype(float)
		self.df['Date'] = self.df['Date'].apply(lambda x: x.split(' ')[0])
		self.df['Date'] = pd.to_datetime(self.df['Date'], format= '%Y-%m-%d')
		
	def download_database(self):
		self.gdata()				
		filetypes = [('excel', '*.xlsx')]

		path = filedialog.asksaveasfilename(filetypes=filetypes, defaultextension=filetypes, initialfile='database_file')
		self.df.to_excel(path, index=False)

	def check(self):
		self.gdata()
		# merge Cash Journal with Gsheet, return only matching Receipt No (INNER JOIN)
		df_merge = pd.merge(app_inquiry.df_result, self.df, left_on=['Receipt No'], right_on=['Receipt'])

		df_result_row = app_inquiry.df_result.shape[0]
		per_complete = f'{((df_merge.shape[0]/df_result_row)*100):.2f}%'
		
		messagebox.showinfo(title='% Complete', message=f'{per_complete} received WHT')

	def pending(self):
		self.gdata()
		df_pending = pd.merge(app_inquiry.df_result, self.df, left_on='Receipt No', right_on='Receipt',how='outer', indicator=True)
		df_pending = df_pending[df_pending['_merge']=='left_only']

		filetypes = [('excel', '*.xlsx')]

		path = filedialog.asksaveasfilename(filetypes=filetypes, defaultextension=filetypes, initialfile='pending WHT')
		df_pending.to_excel(path, index=False)
		print(df_pending.shape)
		
def main():		
	app_inquiry.draw()

	root.bind('<Escape>', lambda x: root.quit())	
	root.mainloop()

app_inquiry = App_Inquiry(root)
app_wht = App_WHT(root)

if __name__ == '__main__':
	main()
