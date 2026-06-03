# Deployment Guide

Aplikasi BCA Statement Extractor ini didesain agar mudah dideploy di berbagai _environment_, terutama dengan menggunakan Docker.

## Prasyarat
- Docker terinstal di server staging/production Anda.
- Docker Compose terinstal (biasanya sudah include pada Docker versi terbaru).
- Port `8501` terbuka di firewall server Anda.

## Langkah-langkah Deploy

1. **Clone Repository (atau copy seluruh folder project ini)** ke server Anda.
   ```bash
   git clone <url-repo-anda>
   cd mutasi_bca
   ```

2. **Jalankan Docker Compose**
   Gunakan perintah berikut untuk membangun image dan menjalankan container di background (detached mode):
   ```bash
   docker-compose up -d --build
   ```

3. **Verifikasi**
   Akses aplikasi melalui browser:
   ```
   http://<ip-server-anda>:8501
   ```

## Konfigurasi Lanjutan
Environment variables yang didukung untuk mengatur limitasi atau konfigurasi aplikasi dapat diubah di dalam file `docker-compose.yml`:
- `LOG_LEVEL`: Mengubah level log (DEBUG, INFO, WARNING, ERROR). Default `INFO`.
- `MAX_PDF_SIZE`: Mengatur batas maksimal ukuran PDF yang diizinkan untuk diproses (dalam bytes). Default `104857600` (100MB).

## Maintenance

- **Melihat Log Aplikasi**:
  ```bash
  docker-compose logs -f
  ```
- **Menghentikan Aplikasi**:
  ```bash
  docker-compose down
  ```
- **Restart Aplikasi**:
  ```bash
  docker-compose restart
  ```

> [!TIP]
> Jika server Anda sudah di-*proxy* menggunakan Nginx/Apache, Anda bisa mem-forward HTTP request menuju port `8501`. Pastikan Anda mengatur WebSocket Support di proxy Anda karena Streamlit membutuhkan koneksi WebSockets aktif.
