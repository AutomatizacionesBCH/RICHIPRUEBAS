FROM nginx:alpine
COPY dashboard_sii.html /usr/share/nginx/html/index.html
COPY Logo/ /usr/share/nginx/html/Logo/
EXPOSE 80
