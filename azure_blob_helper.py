from azure.storage.blob import BlobServiceClient
import os

# ==========================================================
# AZURE STORAGE CONFIGURATION
# ==========================================================

try:

    from local_settings import (
        AZURE_STORAGE_CONNECTION_STRING
    )

    CONNECTION_STRING = (
        AZURE_STORAGE_CONNECTION_STRING
    )

except Exception:

    CONNECTION_STRING = None

CONTAINER_NAME = "swiics-adls-data-dev"

if not CONNECTION_STRING:

    raise ValueError(
        "AZURE_STORAGE_CONNECTION_STRING not found."
    )

blob_service = BlobServiceClient.from_connection_string(
    CONNECTION_STRING
)

# ==========================================================
# DOWNLOAD BLOB
# ==========================================================

def download_blob(
    blob_path,
    local_file
):

    blob_client = blob_service.get_blob_client(
        container=CONTAINER_NAME,
        blob=blob_path
    )

    os.makedirs(
        os.path.dirname(local_file),
        exist_ok=True
    )

    with open(
        local_file,
        "wb"
    ) as file:

        file.write(
            blob_client.download_blob().readall()
        )

    print(
        f"Downloaded: {blob_path}"
    )

# ==========================================================
# UPLOAD BLOB
# ==========================================================

def upload_blob(
    local_file,
    blob_path
):

    blob_client = blob_service.get_blob_client(
        container=CONTAINER_NAME,
        blob=blob_path
    )

    with open(
        local_file,
        "rb"
    ) as data:

        blob_client.upload_blob(
            data,
            overwrite=True
        )

    print(
        f"Uploaded: {blob_path}"
    )