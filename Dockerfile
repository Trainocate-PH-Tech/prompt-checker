FROM python:3.14-slim

RUN useradd --user-group --no-create-home python && \
    mkdir -p /app /home/python && \
    chown -R python:python /app /home/python

USER python
WORKDIR /app

COPY requirements.txt .

RUN pip install --user --no-cache-dir -r requirements.txt && \
    mkdir input-files output-files

COPY app.py copilot_rubrics.pdf ./

ENTRYPOINT ["python", "app.py"]
CMD ["--help"]