#######################
#    For AWS Server   #
#######################

# Line Dev 建立一支 Bot＋設定完成
# AWS 防火牆（安全群組）Port 開好
# AWS 彈性 IP 綁定

sudo apt update
sudo apt install certbot python3-certbot-nginx

sudo vim /etc/nginx/sites-available/default
"""
server {
    listen 80;
    server_name memom.ddns.net;

    location / {
        proxy_pass http://127.0.0.1:5000;  # Flask 默認運行在 5000 端口
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
"""

python3 -m venv doc_venv
source doc_venv/bin/activate
pip install -r requirements.txt

nohup python3 main_p2.py &
sudo certbot --nginx -d memom.ddns.net

# 測試自動刷新憑證
sudo certbot renew --dry-run


#######################
#    For GCP Server   #
#######################

# [First] 使用 noip 註冊一個 Domain + IPv4 Address 綁定正確
sudo apt-get install python3 python3-pip nginx
sudo apt-get install certbot python3-certbot-nginx
sudo pkill -f nginx
sudo ufw allow 80
sudo ufw allow 443
sudo ufw reload
sudo certbot certonly --standalone -d your_domain
sudo openssl dhparam -out /etc/letsencrypt/ssl-dhparams.pem 2048

sudo vim /etc/letsencrypt/options-ssl-nginx.conf
"""
ssl_session_cache shared:le_nginx_SSL:1m;
ssl_session_timeout 1440m;

ssl_protocols TLSv1.2 TLSv1.3;
ssl_prefer_server_ciphers on;
ssl_ciphers "ECDHE-ECDSA-CHACHA20-POLY1305:ECDHE-RSA-CHACHA20-POLY1305:ECDHE-ECDSA-AES128-GCM-SHA256:ECDHE-RSA-AES128-GCM-SHA256:ECDHE-ECDSA-AES256-GCM-SHA384:ECDHE-RSA-AES256-GCM-SHA384:DHE-RSA-AES128-GCM-SHA256:DHE-RSA-AES256-GCM-SHA384:ECDHE-ECDSA-AES128-SHA256:ECDHE-RSA-AES128-SHA256:ECDHE-ECDSA-AES128-SHA:ECDHE-RSA-AES128-SHA:ECDHE-ECDSA-AES256-SHA384:ECDHE-RSA-AES256-SHA384:ECDHE-ECDSA-AES256-SHA:ECDHE-RSA-AES256-SHA:DHE-RSA-AES128-SHA256:DHE-RSA-AES256-SHA256:DHE-RSA-AES128-SHA:DHE-RSA-AES256-SHA:ECDHE-ECDSA-DES-CBC3-SHA:ECDHE-RSA-DES-CBC3-SHA:EDH-RSA-DES-CBC3-SHA:AES128-GCM-SHA256:AES256-GCM-SHA384:AES128-SHA256:AES256-SHA256:AES128-SHA:AES256-SHA:DES-CBC3-SHA:!DSS";
ssl_ecdh_curve secp384r1; # Requires nginx >= 1.1.0
ssl_session_tickets off; # Requires nginx >= 1.5.9
ssl_stapling on; # Requires nginx >= 1.3.7
ssl_stapling_verify on; # Requires nginx => 1.3.7
resolver 8.8.8.8 8.8.4.4 valid=300s;
resolver_timeout 5s;
"""

sudo vim /etc/nginx/sites-available/default
"""
server {
    listen 80;
    server_name your_domain;

    location / {
        return 301 https://$host$request_uri;
    }
}

server {
    listen 443 ssl;
    server_name your_domain;

    ssl_certificate /etc/letsencrypt/live/your_domain/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/your_domain/privkey.pem;
    include /etc/letsencrypt/options-ssl-nginx.conf;
    ssl_dhparam /etc/letsencrypt/ssl-dhparams.pem;

    location / {
        proxy_pass http://127.0.0.1:5000;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }
}
"""

sudo pkill -f nginx
sudo certbot --nginx
