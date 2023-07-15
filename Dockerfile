FROM python:3-alpine

ARG USERNAME=worker

RUN pip install --upgrade pip

RUN adduser -D $USERNAME
USER $USERNAME

WORKDIR /home/$USERNAME

COPY --chown=$USERNAME:$USERNAME requirements.txt ./
RUN pip install --user --no-cache-dir -r requirements.txt
RUN rm requirements.txt

ENV PATH="/home/${USERNAME}/.local/bin:${PATH}"

COPY --chown=$USERNAME:$USERNAME ./app ./app

WORKDIR /home/$USERNAME/app

CMD [ "python", "-u", "./app.py" ]
