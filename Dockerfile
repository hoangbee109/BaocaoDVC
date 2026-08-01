FROM node:trixie-slim

WORKDIR /app

COPY package*.json ./

RUN npm install

RUN mkdir -p uploads

COPY . .

EXPOSE 3000

CMD ["node", "server.js"]