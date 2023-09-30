# Use the official lightweight Python image.
FROM python:3.11-slim

# Set the working directory
WORKDIR /app

# Copy the requirements.txt and install the Python dependencies.
COPY requirements.txt requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copy the script into the container.
COPY img_to_excel.py img_to_excel.py

# Command to run on container start. 
# Here, we'll simply execute the Python script.
CMD ["python", "./img_to_excel.py"]