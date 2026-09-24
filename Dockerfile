FROM node:22-alpine AS build
WORKDIR /usr/local/app
ARG VITE_BASE=/vue-sdb-rec-sign/
ENV VITE_BASE=${VITE_BASE}
COPY package.json package-lock.json* ./
RUN npm ci
COPY . .
RUN npm run build

FROM nginx:1.27-alpine
LABEL authors="sciencedb"
EXPOSE 9093
WORKDIR /usr/share/nginx/html
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY --from=build /usr/local/app/dist /usr/share/nginx/html/vue-sdb-rec-sign
CMD ["nginx", "-g", "daemon off;"]
