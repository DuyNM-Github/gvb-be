FROM --platform=linux/amd64 python:3.10-slim
ENV PYTHONUNBUFFERED=1
WORKDIR /app
COPY . /app
RUN apt-get update
RUN apt-get install -y gcc
RUN apt-get install -y libreoffice-writer
RUN pip install -r requirements.txt
EXPOSE 8000
RUN ["chmod", "+x", "run.sh"]
CMD ["./run.sh"]
