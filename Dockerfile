# ── Stage 1: build the React/Vite app ────────────────────────────────────────
FROM node:20-alpine AS builder

WORKDIR /app

# Install dependencies first (layer-cached unless package.json changes)
COPY package.json ./
RUN npm install

# Copy source and build
COPY . .
RUN npm run build

# ── Stage 2: serve with nginx on port 7860 ────────────────────────────────────
FROM nginx:alpine

# Replace default nginx config with our custom one
COPY nginx.conf /etc/nginx/conf.d/default.conf

# Copy the built app
COPY --from=builder /app/dist /usr/share/nginx/html

EXPOSE 7860

CMD ["nginx", "-g", "daemon off;"]
