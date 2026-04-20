"""Fetcher: talks to remote endpoints to retrieve report lists and files.

Connections:
- Called by the GUI launcher to download data for a date range.
- Saves downloaded reports to latest/ then hands off to Combiner.generate to build record2025.csv.
- Exposes missing_fetch for the missing_generate utility to re-download only missing (branch, POS, date) records.
"""
import random
import os
import time
import requests
import datetime
import os
import pandas
import datetime
import unicodedata
import shutil
from combiner import Combiner
import concurrent.futures

class Receive():
	"""Orchestrates fetching POS data across branches and dates, downloading files, and combining them."""

	def __init__(self, start_time,end_time,datearr=pandas.DataFrame()):		
		"""Initialize with a start/end date range or a custom date DataFrame, and load branch list."""
		self.exitFlag = 0

		self.parentDir = os.getcwd()
		#parentDir = "/storage/emulated/0"
		self.session = requests.Session()

		if datearr.empty:
			self.sfull = start_time.split("-")
			self.efull = end_time.split("-")
			print(self.sfull)
			print(self.efull)
	
			self.start = datetime.date(int(self.sfull[0]),int(self.sfull[1]),int(self.sfull[2]))
			self.end = datetime.date(int(self.efull[0]),int(self.efull[1]),int(self.efull[2]))

			self.dlist = pandas.date_range(self.start,self.end,freq='d')
		else:
			self.dlist = datearr
		self.branches = []
		self.f1=open( self.parentDir + "/settings/branches.txt","r")
		self.file=self.f1.read()
		for row in self.file.splitlines():
			self.branches.append(row)
		self.empty = []

	def _log(self, payload):
		try:
			cb = getattr(self, "log_callback", None)
			if callable(cb):
				cb(payload)
		except Exception:
			pass

	def send(self, branch, pos, date):
		"""Request the report list for a given branch/POS/date from the remote endpoint."""
		filt_date = str(date)[:10]
		print("Arguments")
		print(branch)
		print(pos)
		print(date)
		print("Fetching: " + branch +" POS #" + str(pos) + " for date " + str(filt_date))
		self._log({"kind": "request", "branch": branch, "pos": pos, "date": filt_date, "message": "Requesting file list"})
		for i in range(3):
			try:
				self.headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:55.0) Gecko/20100101 Firefox/55.0',
				'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    			'Accept-Language': 'en-US,en;q=0.5',
    			'Content-Type': 'application/x-www-form-urlencoded',
    			'Referer': 'https://biggsph.com/',
    			'Origin': 'https://biggsph.com'}
				self.url = 'https://biggsph.com/biggsinc_loyalty/controller/fetch_list2.php'
				self.data = {'branch' : branch, 'pos': pos, 'date': filt_date}
				self.r = requests.Request('POST',self.url, data = self.data, headers = self.headers).prepare()
				self.resp = self.session.send(self.r, timeout=30)
				if "<!doctype html>" in self.resp.text:
					return [""]
				else:
					return self.resp.text.split(",")
				break
			except Exception as e:
				print(e)

	def download_file(self, url, destination):
		"""Download a report file by URL into a local destination folder, returning local path or blank."""
		for i in range(3):
			try:
				if(not url == ""):
					self.headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10.12; rv:55.0) Gecko/20100101 Firefox/55.0',}
					self.local_filename = self.parentDir + "/" + destination + "/" + url.split('/')[-1]
					with self.session.get("https://biggsph.com/biggsinc_loyalty/controller/" + url, stream=True, headers = self.headers, timeout=60) as r:
						r.raise_for_status()
						with open(self.local_filename, 'wb') as f:
							for chunk in r.iter_content(chunk_size=65536):
								f.write(chunk)
					return self.local_filename
				else:
					return ""
				break
			except Exception as e:
				print(e)
				return ""

	def process(self,filearray,pos):
		"""Download and stage all files in the given file list for a POS; build latest artifacts."""
		try:
			filetypes = ["rd1800", "blpr", "discount", "rd5000", "rd5500", "rd5800", "rd5900"]
			self.maxfile = {ftype: {"file": "", "count": 0} for ftype in filetypes}
			self.tempfile = ""
			self.maxcount = 0
			self.tempcount = 0
			self.local = ""
			for file in filearray:
				if file != "":
					print("Downloading file: " + file)
					self._log({"kind": "file", "branch": getattr(self, "_current_branch", ""), "pos": pos, "date": str(getattr(self, "_current_date", ""))[:10], "file": file})
					self.tempfile = self.download_file(file,"latest")
					
		except Exception as e:
			print("Process Error")
			print(e)

	def clean(self,directory):
		"""Remove all contents of a directory (files and folders)."""
		for filename in os.listdir(directory):
			file_path = os.path.join(directory, filename)
			try:
				if os.path.isfile(file_path) or os.path.islink(file_path):
					os.unlink(file_path)
				elif os.path.isdir(file_path):
					shutil.rmtree(file_path)
			except Exception as e:
				print('Failed to delete %s. Reason: %s' % (file_path, e))

	def fetch(self):
		"""Main fetch loop: for each date and branch, get POS data, process, then combine and log."""
		self.clean(self.parentDir + '/latest')
		self.clean(self.parentDir + '/temp')
		try:
			total = getattr(self, "progress_total", None)
			done = 0
		except Exception:
			total = None
			done = 0
		for date in self.dlist:
			try:
				if getattr(self, "exitFlag", 0):
					break
			except Exception:
				pass
			print("\nFetching "+ str(date) +" \n")
			for branch in self.branches:
				try:
					if getattr(self, "exitFlag", 0):
						break
				except Exception:
					pass
				self._current_branch = branch
				self._current_date = date
				print("\n\tFetching "+ branch +"\n")
				with concurrent.futures.ThreadPoolExecutor(max_workers=2) as ex:
					f1 = ex.submit(self.send, branch, 1, date)
					f2 = ex.submit(self.send, branch, 2, date)
					try:
						pos1 = f1.result()
					except Exception as e:
						print(e)
						pos1 = []
					try:
						pos2 = f2.result()
					except Exception as e:
						print(e)
						pos2 = []
				self.process(pos1,1)
				self.process(pos2,2)
				try:
					done += 1
					hook = getattr(self, "progress_hook", None)
					if callable(hook):
						hook(done, total if total is not None else done)
				except Exception:
					pass
			compress = Combiner()
			compress.generate()
			self.clean(self.parentDir + '/latest')
			f3 = open(self.parentDir + "/last_record.log","w")
			f3.write((date + datetime.timedelta(days=1)).strftime("%Y-%m-%d"))
			f3.close()
		print ("Maxfiles that have 0 entries:")
		for x in range(len(self.empty)):
			print (self.empty[x])

	# def missing_fetch(self, branches):
	# 	for date in self.dlist.index:
	# 		print("\nFetching "+ str(date) +" \n")
	# 		for branch in self.dlist.columns.values:
	# 			#print(branch)
	# 			if (self.dlist.loc[date,branch] == 0):
	# 				print(branch)
	# 				print("\n\tFetching "+ str(branch) +"\n")
	# 				pos1 = self.send(branch, 1, date)
	# 				pos2 = self.send(branch, 2, date)
	# 				self.process(pos1)
	# 				self.clean(self.parentDir + '/temp')
	# 				self.process(pos2)
	# 				self.clean(self.parentDir + '/temp')
	# 		compress = Combiner()
	# 		compress.generate()
	# 		print("Initiate Append")
	# 		compress.append()
	# 		self.clean(self.parentDir + '/latest')
	# 	print ("Maxfiles that have 0 entries:")
	# 	for x in range(len(self.empty)):
	# 		print (self.empty[x])

	def missing_fetch(self, branches_missing):
		"""Fetch only the missing records for given branches/POS dates structure and re-generate outputs."""
		"""
		Fetch only missing records based on branches_missing structure.
		
		branches_missing example:
		{
			'SMNAG': {
				1: ['2025-08-15', '2025-08-16', ...],
				2: ['2025-08-15', '2025-08-16', ...]
			},
			'OTHER_BRANCH': {
				1: ['2025-08-20'],
				2: ['2025-08-20']
			}
		}
		"""
		# Clean up directories first
		self.clean(self.parentDir + '/latest')
		self.clean(self.parentDir + '/temp')

		for branch, pos_dict in branches_missing.items():
			print(f"\nFetching missing data for branch: {branch}\n")

			for pos, dates in pos_dict.items():
				for date in dates:
					self._current_branch = branch
					self._current_date = date
					print(f"\tFetching branch {branch}, pos {pos}, date {date}")

					result = self.send(branch, pos, date)

					self.process(result, pos)

					# Clean temp after each pos fetch
					# self.clean(self.parentDir + '/temp')

			# After processing a branch, you may want to combine/compress
			compress = Combiner()
			compress.generate()
			self.clean(self.parentDir + '/latest')

			# Update log with the last processed date
			if dates:  # only if there are dates
				last_date = max(dates)
				# with open(self.parentDir + "/last_record.log", "w") as f3:
				# 	f3.write((datetime.datetime.strptime(last_date, "%Y-%m-%d") + datetime.timedelta(days=1)).strftime("%Y-%m-%d"))

		print("Maxfiles that have 0 entries:")
		for x in range(len(self.empty)):
			print(self.empty[x])
	def missing_pos_fetch(self):
		"""Re-fetch missing POS rows based on a table layout (date x branch/POS) and append results."""
		for date in self.dlist.index:
			print("\nFetching "+ str(date) +" \n")
			for branch in self.dlist.columns.values:
				#print(branch)
				if (self.dlist.loc[date,branch] == 0):
					print(branch)
					print("\n\tFetching "+ str(branch) +"\n")
					pos = self.send(branch[:-1], int(branch[-1:]), date)
					print("POS: ")
					print(pos)
					self.process(pos, int(branch[-1:]))
					self.clean(self.parentDir + '/temp')
			compress = Combiner()
			compress.generate()
			print("Initiate Append")
			compress.append()
			self.clean(self.parentDir + '/latest')
		print ("Maxfiles that have 0 entries:")
		for x in range(len(self.empty)):
			print (self.empty[x])
