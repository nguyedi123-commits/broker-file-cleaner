# Broker File Cleaner Web App

This project includes:

- `broker_file_cleaning/process_templates.py`: the shared cleaning engine
- `streamlit_app.py`: the Streamlit upload-and-download app
- `.streamlit/config.toml`: Streamlit upload configuration

## Local app run

Use Python `3.12` for the cleanest install path. The pinned `pandas` build in this project may not have wheels for Python `3.14`.

1. Install dependencies:

   ```powershell
   pip install -r requirements.txt
   ```

2. Start the app:

   ```powershell
   streamlit run streamlit_app.py
   ```

## Local batch mode

The original batch workflow still works:

```powershell
python broker_file_cleaning\process_templates.py
```

It reads from `broker_file_cleaning\Input` and writes cleaned output to `broker_file_cleaning\Output`.

## Streamlit Community Cloud deployment

1. Push this project to GitHub.
2. Sign in to Streamlit Community Cloud.
3. Create a new app from the repository.
4. Set the entrypoint to `streamlit_app.py`.
5. Deploy.

The deployed app is stateless:

- users upload raw files
- files are processed in a temporary directory
- the cleaned CSV is returned for download
- temporary input and output files are deleted automatically

## Energy unit rule

This cleaner only processes interval energy data in `kWh` or `MWh`.

- `kWh` is used as-is
- `MWh` is converted to `kWh`
- reactive or non-energy units such as `kVArh` are ignored

This rule applies to both the local batch workflow and the Streamlit upload flow.

## Upload limits in the app

- Max files per upload: 50
- Max total batch size: 300 MB
- Max single file size: 100 MB

These app-level limits are enforced before processing starts.
