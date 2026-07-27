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

1. Upload this project folder to a GitHub repository.
2. In Render, choose `New` -> `Web Service`.
3. Connect the GitHub repository.
4. Set the runtime to Python.
5. Use this build command:

```bash
pip install -r requirements.txt
```

6. Use this start command:

```bash
HOST=0.0.0.0 python app/server.py
```

Render will provide the `PORT` automatically.

This repository also includes `render.yaml`, so Render may detect the service settings automatically.

## Why GitHub Pages Shows A 404 Or Static Page

GitHub Pages only serves static files such as HTML, CSS, and JavaScript. It does not run Python, Pandas, OpenPyXL, file uploads, or the reconciliation server.

The root `index.html` is included only so GitHub Pages does not show a missing-file error. The working upload/reconciliation app must be deployed to a Python hosting service.

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
