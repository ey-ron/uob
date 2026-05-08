# Use the official Playwright image which includes all OS dependencies
FROM mcr.microsoft.com/playwright:v1.59.1-jammy

# Set the working directory
WORKDIR /app

# Copy package files and install dependencies
COPY package*.json ./
RUN npm install

# Copy the rest of your application code
COPY . .

# Expose the port your Express server uses
EXPOSE 3000

# Start the application
CMD ["node", "index.js"]