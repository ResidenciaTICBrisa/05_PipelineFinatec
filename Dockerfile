FROM python:3.10.12

ENV PYTHONUNBUFFERED=1

WORKDIR /code


COPY requirements.txt .

RUN pip install --upgrade pip

RUN pip install -r requirements.txt

COPY . .

EXPOSE 8000

CMD ["python3","manage.py","runserver","0.0.0.0:8000"]