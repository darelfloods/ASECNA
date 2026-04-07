FROM node:20-slim

WORKDIR /app

# Install build dependencies for better-sqlite3
RUN apt-get update && apt-get install -y python3 make g++ && rm -rf /var/lib/apt/lists/*

# Copy package files and install ALL dependencies (dev needed for build)
COPY package.json package-lock.json ./
RUN npm ci

# Copy source files
COPY server/ ./server/
COPY public/ ./public/
COPY src/ ./src/
COPY index.html ./
COPY vite.config.ts ./
COPY tsconfig.json ./

# Build frontend
RUN npm run build

# Remove dev dependencies
RUN npm prune --omit=dev

# Create data directory for SQLite persistence
RUN mkdir -p /data

ENV NODE_ENV=production
ENV PORT=8080

EXPOSE 8080

CMD ["node", "server/standalone-commonjs.js"]
