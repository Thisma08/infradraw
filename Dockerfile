FROM python:3.9-slim
WORKDIR /app
COPY insertion.py /app/
COPY dessin.py /app/
COPY config.ini /app/
RUN pip install requests mysql-connector-python pywin32 configparser
CMD ["python", "insertion.py", "dessin.py"]
