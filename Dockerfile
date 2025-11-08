# Use Miniconda base image
FROM continuumio/miniconda3

# Set working directory
WORKDIR /app

# Create and activate a new environment
RUN conda create -y -n rebalance_env python=3.11 \
    && echo "conda activate rebalance_env" >> ~/.bashrc

# Install main deps with conda
RUN conda install -n rebalance_env -y \
    pandas \
    numpy \
    openpyxl \
    xlsxwriter \
    requests \
    && conda clean -afy

# Copy requirements first for caching
COPY requirements.txt /app/

# Install pip-only packages
RUN conda run -n rebalance_env pip install --no-cache-dir -r requirements.txt

# ✅ Copy application code (HTML, Python, static files)
COPY . /app/

# Expose port (Render uses dynamic $PORT)
EXPOSE 8000

# ✅ CMD must always be last
CMD conda run --no-capture-output -n rebalance_env uvicorn app:app --host 0.0.0.0 --port $PORT
