# Use Miniconda base image
FROM continuumio/miniconda3

# Set working directory
WORKDIR /app

# Create and activate a new environment
RUN conda create -y -n rebalance_env python=3.11 \
    && echo "conda activate rebalance_env" >> ~/.bashrc

# Install main deps with conda (faster, prebuilt binaries)
RUN conda install -n rebalance_env -y \
    pandas \
    numpy \
    openpyxl \
    xlsxwriter \
    requests \
    && conda clean -afy

# Copy requirements first (better caching)
COPY requirements.txt /app/

# Install pip-only packages inside the conda env (fastapi, uvicorn, yfinance, ta, etc.)
RUN conda run -n rebalance_env pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . /app/

# Expose port
EXPOSE 8000

# Default command: run FastAPI app inside conda env
CMD ["conda", "run", "--no-capture-output", "-n", "rebalance_env", "uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
