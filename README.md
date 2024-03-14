# HOW TO RUN
The project is already pre-configured with the necessary settings to work with any environment.
To run locally, proceed to the project directory.
```
python manage.py runserver
```
To run via container, the project also has pre-configured Dockerfile. Proceed to build an image and run the container. The project listens on port 8000
```
docker build -t <your-prefered-image-name>
docker run -p 8000:8000 --name <container-name> <image-name>
```
