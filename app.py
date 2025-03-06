from flask import Flask, render_template, request
import os
from azure.storage.blob import BlobServiceClient
from office365.graph_client import GraphClient

app = Flask(__name__)

# Azure Storage Config
AZURE_CONNECTION_STRING = "DefaultEndpointsProtocol=https;AccountName=wordformstorage;AccountKey=lgI53sN3ubrhptYW1bTbhevcCgdRf8C77Dh3Uz/g5WWU5Sx91S4RxPEU+DhZgwxqLnNN41VWDsEr+AStMqCSEw==;EndpointSuffix=core.windows.net"
CONTAINER_NAME = "submitted-pdfs"
blob_service_client = BlobServiceClient.from_connection_string(AZURE_CONNECTION_STRING)

# Microsoft Graph API Config
GRAPH_CLIENT_ID = "1a704878-8f32-4847-82c9-33172216189f"
GRAPH_CLIENT_SECRET = "ojk8Q~hNuMdCQswZoeLjhN5t9chXe8uJED8mza-8"
TENANT_ID = "7a685b45-ee57-40a4-8603-4469227a010e"

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/form", methods=["GET"])
def open_form():
    word_file_url = "https://wordformstorage.blob.core.windows.net/form-files/form.docx"
    return render_template("form.html", word_file_url=word_file_url)

@app.route("/convert", methods=["POST"])
def convert_to_pdf():
    word_file_url = request.form.get("doc_url")

    # Convert Word to PDF using Microsoft Graph API
    graph_client = GraphClient(GRAPH_CLIENT_ID, GRAPH_CLIENT_SECRET, TENANT_ID)
    pdf_data = graph_client.convert_to_pdf(word_file_url)

    # Save the PDF in Azure Blob Storage
    pdf_filename = "submitted_form.pdf"
    blob_client = blob_service_client.get_blob_client(container=CONTAINER_NAME, blob=pdf_filename)
    blob_client.upload_blob(pdf_data, overwrite=True)

    return "Form submitted! PDF has been saved to Azure."

if __name__ == "__main__":
    app.run(debug=True)
