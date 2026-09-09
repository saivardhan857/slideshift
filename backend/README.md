# SlideShift — Backend

FastAPI + python-pptx service that transfers the content of a source `.pptx`
into a college template `.pptx`, preserving text, images and tables.

## Run locally

```
cd backend
pip install -r requirements.txt
uvicorn main:app --host 0.0.0.0 --port 8765
```

The frontend (`../frontend/index.html`) is served at `/`.

## Key modules

| File | Responsibility |
| --- | --- |
| `parser.py` | Parse a source PPTX into structured `ParsedSlide` data |
| `layout_matcher.py` | Classify each slide and pick the best template layout |
| `transfer.py` | Write parsed content into the template and save the output |
| `validator.py` | Sanity-check the generated PPTX |
| `main.py` | FastAPI endpoints (`/api/transfer`, `/api/download/{id}`, `/health`) |

## Tests

```
python test_transfer.py
python test_qa_regression.py
python test_title_detection.py
```

## Deployment

Deployed on Render from `render.yaml` (branch `main`, root directory `backend`,
auto-deploy on commit). Production: https://slideshift.onrender.com
