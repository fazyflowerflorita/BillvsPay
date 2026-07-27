# Website Deployment

This project is now deployable as a small Python web app.

## What Changed

The app can still run locally, but it also supports website hosting through environment variables:

- `HOST`: use `0.0.0.0` on a hosted server.
- `PORT`: supplied automatically by most hosting providers.

## Recommended Hosting Options

Use one of these:

- Azure App Service if this needs to stay inside Microsoft/company infrastructure.
- Render, Railway, or similar app hosting for a quick public web URL.
- An internal Windows or Linux server if files should not leave the company network.

## Render-Style Deployment

1. Upload this project folder to a Git repository.
2. Create a new Web Service.
3. Set the runtime to Python.
4. Use this build command:

```bash
pip install -r requirements.txt
```

5. Use this start command:

```bash
HOST=0.0.0.0 python app/server.py
```

Render will provide the `PORT` automatically.

## Azure App Service Deployment

Use these settings:

- Runtime: Python 3.12
- Startup command:

```bash
python app/server.py
```

Set app settings:

```text
HOST=0.0.0.0
```

Azure supplies the port through its hosting layer.

## Important Notes

- Uploaded files are saved in `uploads/`.
- Generated reconciliation reports are saved in `output/`.
- On cloud platforms with temporary storage, old reports may be cleared when the service restarts.
- For production use, add authentication before exposing this outside the company network.
- If reports contain employee/payroll data, prefer Azure App Service or an internal server over a public demo host.

