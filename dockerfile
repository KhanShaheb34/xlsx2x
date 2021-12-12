FROM alpine:3

RUN apk add --update python3 python3-dev py-pip build-base openjdk8-jre \
  && pip install virtualenv \
  && rm -rf /var/cache/apk/*
RUN pip install aspose.cells bs4 flask waitress weasyprint opencv-python

WORKDIR /app
EXPOSE 5000

COPY '*.py' '.'
CMD [ "python3", "app.py" ]