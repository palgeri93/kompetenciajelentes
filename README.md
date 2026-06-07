# Kompetenciamérés PDF generátor

Webes Streamlit alkalmazás, amely Excel feltöltéséből PDF diagramokat készít.

## Helyi futtatás

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
streamlit run app.py
```

## GitHub + Streamlit Cloud publikálás

1. Hozz létre egy új GitHub repository-t.
2. Töltsd fel ezeket a fájlokat:
   - `app.py`
   - `requirements.txt`
   - `.streamlit/config.toml`
   - `README.md`
3. Nyisd meg a Streamlit Community Cloudot.
4. Válaszd ki a GitHub repository-t.
5. Main file path: `app.py`.
6. Deploy.

## Elvárt Excel-szerkezet

Alapértelmezetten a program a `Munka1` munkalap 3. sorát használja fejlécként.
A mintafájlhoz igazítva a következő oszlopok kellenek:

- `NEVEK` vagy `Név`
- `Képességszint`
- `Képességszint.1`
- `Képességpont változás`

