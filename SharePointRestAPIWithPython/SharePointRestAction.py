import requests
import json
import SharePointRestAPISettings
import uuid
import os

# This file containt actions to perform with the Sharepoint Rest API.

# getToken takes your settings from SharePointRestAPISettings and send a requests, the response returns an AccessToken.
# This Function must be called before every other request.


def getToken(settings):
    body = {
        "grant_type": "client_credentials",
        "client_id": settings["appReg_clientId"] + "@" + settings["realm"],
        "client_secret": settings["appReg_clientSecret"],
        "resource": settings["principal"] + "/" + settings["targetHost"] + "@" + settings["realm"]

    }

    response = requests.post("https://accounts.accesscontrol.windows.net/" +
                             settings["realm"] + "/tokens/OAuth/2", data=body).json()

    settings["appReg_bearerToken"] = response["access_token"]


# read_in_chunks split a large file into chuncks for upload with the uploadFile Function if the file its larger than 10mb.
def read_in_chunks(file_object, chunk_size):
    while True:
        data = file_object.read(chunk_size)
        if not data:
            break
        yield data


# uploadFile send a request for uploading a file to SharePoint
# If the file its larger than 10mb, the file its split into multiple parts and send with the /startupload, /continueupload and /finishupload endpoints
def uploadFile(settings, fileName, relativeURL, siteURL):
    offset = 0
    chunckSize = 3 * 1024 * 1024
    header = {
        "Authorization": "Bearer " + settings["appReg_bearerToken"],
        "Accept": "application/json; odata=verbose",
    }

    with open("./" + fileName, "rb") as file:
        fileLength = os.path.getsize("./" + fileName)
        if (fileLength <= chunckSize):
            response = requests.post(
                "https://" + settings["targetHost"] + siteURL + "/_api/web/GetFolderByServerRelativeUrl('" + relativeURL + "')/Files/add(url='" + fileName + "', overwrite=true)", headers=header, data=file).json()

            print(response)
        else:
            response = requests.post(
                "https://" + settings["targetHost"] + siteURL +
                "/_api/web/GetFolderByServerRelativeUrl('" +
                relativeURL + "')/Files/add(url='"
                + fileName + "', overwrite=true)", headers=header).json()

            print(response)
            uploadID = response['d']['UniqueId']

            first = True
            totalBytesRead = 0
            for chunk in read_in_chunks(file, chunckSize):
                totalBytesRead = totalBytesRead + len(chunk)
                if (first):
                    first = False
                    r = requests.post("https://" + settings["targetHost"] + siteURL + "/_api/web/getfilebyserverrelativeurl('" + siteURL + "/" + relativeURL + "/" + fileName + "')/startupload(uploadId=guid'" +
                                      str(uploadID) + "')", headers=header, data=(chunk)).json()
                    print(r)

                elif (totalBytesRead == fileLength):
                    r = requests.post("https://" + settings["targetHost"] + siteURL + "/_api/web/getfilebyserverrelativeurl('" + siteURL + "/" + relativeURL + "/" + fileName + "')/finishupload(uploadId=guid'" +
                                      str(uploadID) + "', fileOffset=" + str(offset) + ")", headers=header, data=(chunk)).json()
                    print(r)

                else:
                    r = requests.post("https://" + settings["targetHost"] + siteURL + "/_api/web/getfilebyserverrelativeurl('" + siteURL + "/" + relativeURL + "/" + fileName + "')/continueupload(uploadId=guid'" +
                                      str(uploadID) + "', fileOffset=" + str(offset) + ")", headers=header, data=(chunk)).json()
                    print(r)

                offset += len(chunk)
                print("%" + "{:.2f}".format(offset *
                                            100 / fileLength) + " Completed")


# getFile send a request to the SharePoint Rest API and retrieve a specific File inside a Folder.
def getFile(settings, fileName, relativeURL, siteURL):
    header = {
        "Authorization": "Bearer " + settings["appReg_bearerToken"],
        "Accept": "application/json; odata=verbose",
    }

    response = requests.get("https://" + settings["targetHost"] + siteURL + "/_api/web/GetFolderByServerRelativeUrl('" +
                            relativeURL + "')/Files('" + fileName + "')/", headers=header).json()

    print(response)


# deleteFile send a request to the SharePoint Rest API and delete a specific File inside a Folder.
def deleteFile(settings, fileName, relativeURL, siteURL):
    header = {
        "Authorization": "Bearer " + settings["appReg_bearerToken"],
        "Accept": "application/json; odata=verbose",
    }

    response = requests.delete("https://" + settings["targetHost"] + siteURL + "/_api/web/GetFolderByServerRelativeUrl('" +
                               relativeURL + "')/Files('" + fileName + "')/", headers=header).json()

    print(response)


# getAllFiles send a request to the SharePoint Rest API and retrieve all the Files inside a Folder.
def getAllFiles(settings, relativeURL, siteURL):
    header = {
        "Authorization": "Bearer " + settings["appReg_bearerToken"],
        "Accept": "application/json; odata=verbose",
    }

    response = requests.get("https://" + settings["targetHost"] + siteURL + "/_api/web/GetFolderByServerRelativeUrl('" +
                            relativeURL + "')/Files", headers=header).json()

    print(response)


# addFolder send a request to create a Folder.
def addFolder(settings, relativeURL, siteURL, folderName):

    header = {
        "Authorization": "Bearer " + settings["appReg_bearerToken"],
        "Accept": "application/json; odata=verbose",
        "Content-Type": "application/json;odata=verbose"
    }

    response = requests.post(
        "https://" + settings["targetHost"] + siteURL + "/_api/Web/Folders/add('" + relativeURL + "/" + folderName + "')", headers=header).json()

    print(response)


# Example of usage(getToken)
getToken(SharePointRestAPISettings.settings)

file_name = "Example.txt"
relative_url = "Shared Documents/General"
site_url = "/sites/ContosoGaming"

# Example of usage (getFile)
getFile(SharePointRestAPISettings.settings, file_name, relative_url, site_url)

Example of usage(getAllFiles)
getAllFiles(SharePointRestAPISettings.settings, relative_url, site_url)

file_to_upload = "ExampleToUpload.jpg"  # File size less than 10mb.

# Example of usage (uploadFile) File size less than 10mb
uploadFile(SharePointRestAPISettings.settings,
           file_to_upload, relative_url, site_url)

large_file_to_upload = "ExampleLargeFile.mp4"  # File size larger than 10mb.

# Example of usage (uploadFile) File size larger than 10mb
uploadFile(SharePointRestAPISettings.settings,
           large_file_to_upload, relative_url, site_url)

folder_name = "NewExampleFolder"

# Example of usage (addFolder)
addFolder(SharePointRestAPISettings.settings,
          relative_url, site_url, folder_name)
