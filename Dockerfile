## Infos https://runnable.com/docker/python/dockerize-your-python-application


# Image de base
FROM python:3

# Ajout des fihiers dans l'image
# Utiliser ADD que si vraiment nécessaire
# ref. https://stackoverflow.com/a/24958548
COPY client.py /
COPY requirements.txt /

# Commandes à executer dans l'immage pour setup le tout
RUN pip3 install -r requirements.txt

# Commande pour executer le script
CMD [ "python", "./client.py" ]
