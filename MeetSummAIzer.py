import os
import time
import logging
import json
from docx import Document
from dotenv import load_dotenv
from openai import AzureOpenAI

# Load environment variables from .env file
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Initialize Azure OpenAI Client
def initialize_azure_openai():
    try:
        client = AzureOpenAI(
            api_key=os.getenv("AZ_OPENAI_API_KEY"),
            api_version="2024-02-01",
            azure_endpoint=os.getenv("AZ_OPENAI_ENDPOINT")
        )
        return client
    except Exception as e:
        logging.error(f"Failed to initialize Azure OpenAI client: {e}")
        exit(1)

# Read and chunk document
def read_and_chunk_document(file_path, chunk_size=15000):
    """
    Reads a document and splits it into manageable chunks.

    Args:
        file_path (str): Path to the document file.
        chunk_size (int): Maximum size of each chunk.

    Returns:
        list: List of text chunks.
    """
    try:
        if file_path.endswith('.docx'):
            doc = Document(file_path)
            full_text = [para.text for para in doc.paragraphs]
            text = "\n".join(full_text)
        elif file_path.endswith('.txt'):
            with open(file_path, 'r') as file:
                text = file.read()
        else:
            logging.error("Unsupported file format. Only .docx and .txt are supported.")
            return []

        chunks = [text[i:i+chunk_size] for i in range(0, len(text), chunk_size)]
        return chunks
    except Exception as e:
        logging.error(f"Error reading or chunking document: {e}")
        return []

# Generate summaries for each chunk
def summarize_chunks(chunks, client, deployment_name, prompt):
    """
    Summarizes each chunk of text using Azure OpenAI.

    Args:
        chunks (list): List of text chunks.
        client (AzureOpenAI): Azure OpenAI client instance.
        deployment_name (str): Deployment name for the model.
        prompt (list): Initial prompt for summarization.

    Returns:
        list: List of summarized responses.
    """
    summaries = []
    for chunk in chunks:
        try:
            messages = prompt + [{"role": "user", "content": chunk}]
            response = client.chat.completions.create(
                model=deployment_name,
                messages=messages,
                max_tokens=4096,
                temperature=0.5,
                top_p=0.5
            )
            if response.choices:
                summaries.append(response.choices[0].message.content.strip())
            else:
                logging.warning(f"No summary generated for chunk: {chunk}")
        except Exception as e:
            logging.error(f"Error summarizing chunk: {e}")
    return summaries

# Main function
def main():
    start_time = time.time()

    # Load environment variables
    api_key = os.getenv("AZ_OPENAI_API_KEY")
    endpoint = os.getenv("AZ_OPENAI_ENDPOINT")
    deployment_name = os.getenv("AZ_OPENAI_DEPLOYMENT_NAME")
    file_path = os.getenv('FILE_PATH')

    if not all([api_key, endpoint, deployment_name, file_path]):
        logging.error("Missing required environment variables.")
        return

    # Load prompts from JSON
    try:
        with open("prompts.json", "r") as file:
            prompts = json.load(file)
    except Exception as e:
        logging.error(f"Error loading prompts: {e}")
        return

    initial_prompt = prompts.get("initial_prompt", [])
    final_prompt = prompts.get("final_prompt", [])

    # Read and chunk the document
    logging.info(f"Processing document: {file_path}")
    chunks = read_and_chunk_document(file_path)
    if not chunks:
        logging.error("No chunks to process.")
        return

    # Initialize Azure OpenAI client
    client = initialize_azure_openai()

    # Summarize each chunk
    logging.info("Summarizing document chunks...")
    chunk_summaries = summarize_chunks(chunks, client, deployment_name, initial_prompt)

    # Combine summaries and generate final summary
    combined_summary = "\n".join(chunk_summaries)
    final_summary = summarize_chunks([combined_summary], client, deployment_name, final_prompt)

    # Save final summary to file
    if final_summary:
        input_filename = os.path.splitext(os.path.basename(file_path))[0]  # Extract filename without extension
        output_file = f"{input_filename}_Summary.txt"
        with open(output_file, "w") as file:
            file.write(final_summary[0])
        logging.info(f"Final summary saved to {output_file}")
    else:
        logging.warning("No final summary generated.")

    # Log execution time
    end_time = time.time()
    logging.info(f"Execution completed in {end_time - start_time:.2f} seconds.")

if __name__ == "__main__":
    main()