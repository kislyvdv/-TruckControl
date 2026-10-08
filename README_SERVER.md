# TruckControl: серверная база

Серверная часть использует SQLite в постоянном Docker volume. API:

- `GET /api/health`
- `GET /api/state`
- `PUT /api/state`

Авторизация: заголовок `X-API-Key`. Для продакшена замените `change-this-key` на длинный секрет и используйте HTTPS.

## Запуск

```bash
docker compose up -d --build
```

После запуска API доступен на `http://SERVER:8080`.

Данные лежат в Docker volume `truckcontrol_data`, поэтому перезапуск контейнера их не удаляет.
