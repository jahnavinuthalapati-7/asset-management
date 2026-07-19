# Asset Management App

Simple Streamlit + Supabase app to track assets across locations — who has each asset, department, location, designation, when issued, and full reallocation history.

Supports 4 asset types: Laptop, i Pad, Desktop, Phone.

## Roles

- **Admin** — sees inventory, can add/reallocate assets, does NOT see reallocation dates or history.
- **Head Admin** — sees everything, including full reallocation history and can create new user accounts.

##  running locally

1. Install Python 3.9+
2. Install dependencies:
3. Update `app.py` with your Supabase credentials 
   - `SUPABASE_URL` = your project URL
   - `SUPABASE_KEY` = your anon key from Supabase Settings → API

4. Run:

## Supabase Setup

1. Create account at [supabase.com](https://supabase.com)
2. Create new project (Region: India/Mumbai)
3. Run SQL queries  (creates tables and default users)
4. Copy Project URL and Anon Key into app.py

## Deploy to internet

Push to GitHub, then deploy free on [Streamlit Cloud](https://streamlit.io/cloud):
- Connect your GitHub repo
- Point to `app.py`
- Get public URL instantly



