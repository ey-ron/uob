# Use official Playwright base image matching v1.59.1
FROM mcr.microsoft.com/playwright:v1.59.1-jammy

WORKDIR /usr/src/app

# Copy package files first
COPY package*.json ./

# Install production dependencies
RUN npm ci --only=production

# Copy application source code
COPY . .

EXPOSE 8080

CMD [ "node", "index.js" ]