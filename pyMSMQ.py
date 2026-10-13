import win32com.client
import win32pdh
import os

MAX_QUEUE_LENGTH = 2000

class MSMQHandler():
	def __init__ (self, queueName, logOutput = False, parent = None):
		super (MSMQHandler, self).__init__ ()
		self._queueNameEnd = queueName
		self._queueName = str(str(os.getenv('COMPUTERNAME')) + "\\private$\\" + str(self._queueNameEnd)).lower()
		self._queueLength = -1
		self._lengthQueryHandle = None
		self._lengthCounterHandle = None
		if(logOutput == True):
			self._logOutput = True
		else:
			self._logOutput = False

		if not self.checkQueueExists(self._queueName):
			self.createQueue(self._queueName)

		if self.checkQueueExists(self._queueName):
			self._queueLength = self.setupMSMQLengthQuery(self._queueName)	# Get initial MSMQ Length
			

	def __del__ (self):
		self.clearMSMQLengthQuery(self._lengthQueryHandle)

	def createQueue(self, queueName):
		try:
			msmq = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")		# Connect to the MSMQ service
			msmq.PathName = queueName									# Set the name of the queue
			msmq.Create()												# Create the queue
			return True
		except Exception as e:
			if self._logOutput:
				print(f"An error occurred [0x00]: {e}")
			return False

	def checkQueueExists(self, queueName):
		try:
			msmq = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")		# Connect to the MSMQ service
			msmq.PathName = queueName									# Set the name of the queue
			msmq.Refresh()												# Try to get information about the queue (this will raise an error if the queue doesn't exist)
			return True													# If no exception, the queue exists
		except Exception as e:
			return False												# If an error occurs, the queue likely doesn't exist

	def setupMSMQLengthQuery(self, queueName):
		try:
			counter_path = f"\\MSMQ Queue({queueName})\\Messages in Queue"									# Define the counter path
			self._lengthQueryHandle = win32pdh.OpenQuery()													# Open the performance counter
			self._lengthCounterHandle = win32pdh.AddCounter(self._lengthQueryHandle, counter_path)			# Add the counter to the query handle
			win32pdh.CollectQueryData(self._lengthQueryHandle)												# Collect the data from the counter
			_, value = win32pdh.GetFormattedCounterValue(self._lengthCounterHandle, win32pdh.PDH_FMT_LONG)	# Get the value of the performance counter
			return value
		except Exception as e:
			if self._logOutput:
				(f"An error occurred [0x01]: {e}")							# Throws an error if queue is of length 0
			return 0
	
	def clearMSMQLengthQuery(self, queryHandle):
		try:
			win32pdh.CloseQuery(queryHandle)							# Close the query and counter handles
			return True
		except Exception as e:
			if self._logOutput:
				print(f"An error occurred [0x02]: {e}")							# Throws an error if queue is of length 0
			return False

	def getMSMQQueueLength(self, counterHandle):
		try:
			win32pdh.CollectQueryData(self._lengthQueryHandle)	
			_, value = win32pdh.GetFormattedCounterValue(counterHandle, win32pdh.PDH_FMT_LONG)		# Get the value of the performance counter
			self._queueLength = value								
			return value
		except Exception as e:
			if self._logOutput:
				print(f"An error occurred [0x03]: {e}")
			return 0

	def sendMessage(self, Title, Body):
		try:
			queueInfo = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
			queueInfo.FormatName = "direct=os:" + self._queueName
			queue = queueInfo.Open(2,0)   										# Open a queue for writing
			msg = win32com.client.Dispatch("MSMQ.MSMQMessage")
			msg.Label =	Title
			msg.Body = Body
			msg.Send(queue)
			queue.Close()

			self.getMSMQQueueLength(self._lengthCounterHandle)

			return self._queueLength
		except Exception as e:
			if self._logOutput:
				print(f"An error occurred [0x04]: {e}")
			return -1
	
	def getMessage(self):
		try:
			queueInfo = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
			queueInfo.FormatName = "direct=os:" + self._queueName
			queue = queueInfo.Open(1,0)   										# Open a queue for reading
			msg=queue.Receive()
			queue.Close()

			self.getMSMQQueueLength(self._lengthCounterHandle)

			return [msg.Body, msg.Label]
		except Exception as e:
			if self._logOutput:
				print(f"An error occurred [0x05]: {e}")
			return -1

	def getQueueLength(self):
		return self._queueLength

	def clearQueue(self):
		try:
			queueInfo = win32com.client.Dispatch("MSMQ.MSMQQueueInfo")
			queueInfo.FormatName = "direct=os:" + self._queueName
			queue = queueInfo.Open(1,0)   										# Open a queue for reading
			queue.Purge()
			queue.Close()

			self.getMSMQQueueLength(self._lengthCounterHandle)

			return self._queueLength
		except Exception as e:
			if self._logOutput:
				print(f"An error occurred [0x05]: {e}")
			return -1