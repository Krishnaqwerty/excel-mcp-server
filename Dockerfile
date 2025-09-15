# Use an official lightweight Python runtime as a parent image
FROM python:3.11-slim

# Set the working directory inside the container
WORKDIR /usr/src/app

# Copy the dependencies file first and install them
# This leverages Docker's layer caching to speed up future builds
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of your application code into the container
COPY app.py .

# Expose port 7777, the standard port for MCP servers
EXPOSE 7777

# The command to run your application when the container starts
CMD ["python", "./app.py"]