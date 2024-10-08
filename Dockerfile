FROM python:3.10
WORKDIR /app
COPY . /app
RUN pip install -r requirements.txt
RUN pip install --user --upgrade pip
CMD ["python", "app.py"]
