# Magpie_DocClassifier 
An alternative method for document chunking and tagging, utilising rule-based algorithms and machine learning for structured content analysis


# Magpie_DocClassifier

## Project Overview
The `Magpie_DocClassifier` is a specialized tool designed for effective document processing across multiple formats including Excel, PDF, DOCX, and PPTX. Unlike traditional document classifiers such as the `GPT-DocClassifier`, `Magpie_DocClassifier` employs a combination of rule-based and AI-driven approaches to optimize document chunking and tagging, ensuring high accuracy and context-sensitive classification.

## Key Features

- **Document Chunking**: Utilizes `doc_decoder_oo.py` to intelligently segment various document types into structured chunks, facilitating detailed analysis and processing.
- **Intelligent Tagging**: Incorporates a two-tier tagging system using OpenAI's GPT 3.5 for initial tag suggestions and a fine-tuned BART model for precise tag assignment, ensuring each chunk is accurately labeled with relevant tags.
- **Cost-Effective AI Use**: Opt for GPT 3.5 to generate initial tag lists, striking a balance between performance and cost, making it suitable for budget-conscious projects.
- **Hierarchical Output**: Outputs chunking and tagging results in a hierarchical JSON format, providing a structured and easy-to-navigate representation of document contents.

## Differences from GPT-DocClassifier

- **Rule-Based Chunking**: Unlike `GPT-DocClassifier` which primarily uses GPT APIs for document chunking, `Magpie_DocClassifier` employs a rule-based method tailored to different document types, enhancing the adaptability to varied document structures.
- **Hybrid Tagging Approach**: Combines AI-driven tag suggestions with a transformer-based classification model, whereas `GPT-DocClassifier` relies solely on GPT models for tagging, providing an added layer of accuracy and context sensitivity in `Magpie_DocClassifier`.
- **Focus on Hierarchical JSON Outputs**: Offers detailed JSON outputs that represent document structures, which is especially useful for subsequent automated processing or manual review, unlike the CSV-centric approach of `GPT-DocClassifier`.

## Getting Started / Process for Executing Code

### Prerequisites
Before you begin, ensure the following prerequisites are met:
- **Python Installation**: Ensure that Python 3.8 or higher is installed on your machine.
- **API Key**: Obtain an API key from OpenAI by visiting their website and registering for access to the GPT 3.5 API. This key must be entered in the `keys.py` file.

### Installation
Install all required Python libraries by running the following command in your project's root directory:
```bash
pip install -r requirements.txt
