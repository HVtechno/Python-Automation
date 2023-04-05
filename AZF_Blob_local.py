#FAST DOWNLOAD FILES FROM BLOB LIBRARIES
from multiprocessing.pool import ThreadPool
#BLOB LIBRARIES
from azure.storage.blob import BlobServiceClient, BlobClient
from azure.storage.blob import ContentSettings, ContainerClient

# Connection string to BLOB storage
MY_CONNECTION_STRING = '(Your Blob connection string)'
# Replace with blob container name
MY_BLOB_CONTAINER = '(Your Blob container repository)'

try:

  class AzureBlobFileDownloader:
    def __init__(self):
      print("Intializing AzureBlobFileDownloader")
 
    # Initialize the connection to Azure storage account
      self.blob_service_client =  BlobServiceClient.from_connection_string(MY_CONNECTION_STRING)
      self.my_container = self.blob_service_client.get_container_client(MY_BLOB_CONTAINER)
    
    def download_all_blobs_in_container(self):
    # get a list of blobs
      my_blobs = self.my_container.list_blobs()
      result = self.run(my_blobs)
      print(result)
 
    def run(self,blobs):
    # Download 10 files at a time!
      with ThreadPool(processes=int(100)) as pool:
        return pool.map(self.save_blob_locally, blobs)
 
    def save_blob_locally(self,blob):
      file_name = blob.name
      filedata = file_name.split('/')[1]
      print(filedata)
      bytes = self.my_container.get_blob_client(blob).download_blob().readall()
      with open(filedata, "wb") as file:
        file.write(bytes)
      return filedata
 
# Initialize class and upload files
  azure_blob_file_downloader = AzureBlobFileDownloader()
  azure_blob_file_downloader.download_all_blobs_in_container()

except:
  pass