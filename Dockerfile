FROM node:20-alpine
WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev
COPY dist/ ./dist/
ENV GOOGLE_SERVICE_ACCOUNT_KEY=""
ENTRYPOINT ["node", "dist/index.js"]
