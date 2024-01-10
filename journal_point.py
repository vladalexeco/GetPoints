class JournalPoint:

	def __init__(self, id, latitude, longitude, search, mad, madError, date, time):
		self.id = id
		self.latitude = latitude
		self.longitude = longitude
		self.search = search
		self.mad = mad
		self.madError = madError
		self.date = date
		self.time = time
	
	def __repr__(self):
		return f"JournalPoint(id = {self.id}"