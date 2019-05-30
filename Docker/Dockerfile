FROM node:10.15.3

RUN npm i -g @pnp/office365-cli
COPY ./loop.sh /usr/src/loop.sh
COPY ./.bashrc /root/.bashrc

CMD /bin/bash /usr/src/loop.sh