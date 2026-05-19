FROM node:18-slim

RUN apt-get update && apt-get install -y libreoffice --no-install-recommends && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .

RUN mkdir -p uploads

EXPOSE 3000
CMD ["node", "index.js"]