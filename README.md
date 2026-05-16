# box_optimizer

Python package and FastAPI service for box optimization workflows.

It includes workbook intake, SKU matching, dimension and weight normalization,
padding, heuristic 3D packing, carton splitting, box standardization, and Excel
report output.

## Quick start

```bash
pip install -r requirements.txt
pytest
```

## API

Run locally:

```bash
uvicorn box_optimizer.api:app --host 0.0.0.0 --port 8000
```

Health check:

```bash
curl http://localhost:8000/health
```

Set `BOX_OPTIMIZER_API_KEY` to require API-key authentication for `POST /optimize`.
Send the key with `X-API-Key` or `Authorization: Bearer ...`.

## Railway

The included Dockerfile is ready for Railway-style deployment. Railway provides
`PORT`; the container starts with:

```bash
uvicorn box_optimizer.api:app --host 0.0.0.0 --port $PORT
```

Copy `.env.example` into Railway variables and set a real `BOX_OPTIMIZER_API_KEY`.

## CLI

```bash
python -m box_optimizer.cli
```
