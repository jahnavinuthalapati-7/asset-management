#  Asset Management App

Simple Streamlit + SQLite app to track asset assignments — who has each asset,
their department/location/designation, when it was issued, and full reallocation history.

## Roles
- **Admin** — sees inventory, can add/reallocate laptops, does NOT see reallocation dates or history log.
- **Head Admin** — sees everything, including full history of every laptop and can create new user accounts.


## How to run locally

1. Install Python 3.9+ if you don't have it.
2. Open a terminal in this folder.
3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```
4. Run the app:
   ```
   streamlit run app.py
   ```
5. It will open automatically in your browser at `http://localhost:8501`.
   The database file `laptops.db` is created automatically in the same folder on first run.

## Making it accessible beyond localhost
- **Same wifi/network:** run `streamlit run app.py --server.address 0.0.0.0` then others on the
  same network can access it via your machine's local IP, e.g. `http://192.168.x.x:8501`.
- **Public internet :** push this folder to a GitHub repo,
  then deploy for free on [Streamlit Community Cloud](https://streamlit.io/cloud) — connect your
  GitHub repo, point it to `app.py`, and it gives you a public URL.

## Git setup (quick)
```
git init
git add .
git commit -m "Initial asset management app"
git branch -M main
git remote add origin <your-repo-url>
git push -u origin main
```
(Streamlit Cloud deploys directly from that repo afterward.)

