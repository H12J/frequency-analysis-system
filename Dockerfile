FROM python:3.9-slim

WORKDIR /app

# Copy requirements first to leverage Docker cache
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all source files
COPY . .

# Create all necessary output directories
RUN mkdir -p extracted_frequencies classified_frequencies reclassified_frequencies

# Set default command to run all scripts in sequence
CMD ["sh", "-c", "python frequency_extractor.py && python frequency_classifier.py && python process_extracted_frequencies.py"]