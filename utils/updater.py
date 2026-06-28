import requests
import os
import sys
import datetime

class GitHubUpdater:
    """
    Compara la fecha del .exe local contra el Latest Release de GitHub.
    Si hay versión nueva, descarga y ejecuta un relevo via .bat.
    """

    def __init__(self, user: str, repo: str):
        self.api_url = (
            f"https://api.github.com/repos/{user}/{repo}/releases/latest"
        )
        self.current_exe = sys.executable
        self.directory = os.path.dirname(self.current_exe)

    # ─── PUBLIC API ───────────────────────────────────────

    def verify(self) -> tuple[bool, str | None, str | None]:
        """
        Returns (has_update, download_url, tag).
        Never raises external exceptions.
        """
        try:
            local_date = self._get_local_exe_date()
            release = self._fetch_latest_release()

            if release is None:
                return False, None, None

            github_date = self._parse_github_date(
                release["published_at"]
            )

            if github_date <= local_date:
                return False, None, None

            url = self._extract_exe_url(release.get("assets", []))
            tag = release.get("tag_name")

            return (url is not None), url, tag

        except Exception as e:
            print(f"[Actualizador] Error verificando: {e}")
            return False, None, None

    def execute_replacement(self, download_url: str):
        """Downloads the new .exe, creates the handover .bat, and exits."""
        temp_path = self._download(download_url)
        bat_path = self._create_handover_bat(temp_path)

        os.startfile(bat_path)
        sys.exit()

    # ─── PRIVATE METHODS ──────────────────────────────────

    def _get_local_exe_date(self) -> datetime.datetime:
        """Modification date of the current executable (UTC)."""
        timestamp = os.path.getmtime(self.current_exe)
        return datetime.datetime.fromtimestamp(
            timestamp, datetime.timezone.utc
        )

    def _fetch_latest_release(self) -> dict | None:
        """Queries the latest release. Returns None if it fails."""
        response = requests.get(self.api_url, timeout=10)
        if response.status_code != 200:
            return None
        return response.json()

    @staticmethod
    def _parse_github_date(date_str: str) -> datetime.datetime:
        """Converts '2025-01-15T10:30:00Z' to UTC datetime."""
        return datetime.datetime.strptime(
            date_str, "%Y-%m-%dT%H:%M:%SZ"
        ).replace(tzinfo=datetime.timezone.utc)

    @staticmethod
    def _extract_exe_url(assets: list) -> str | None:
        """Search for the first asset ending in .exe."""
        return next(
            (a["browser_download_url"] for a in assets
             if a["name"].endswith(".exe")),
            None
        )

    def _download(self, url: str) -> str:
        """
        Downloads to a temporary file.
        Validates that the download was successful before returning.
        """
        temp_path = os.path.join(self.directory, "temp_update.exe")

        response = requests.get(url, stream=True, timeout=120)
        response.raise_for_status()  # ← Raises error if it fails

        downloaded_bytes = 0
        with open(temp_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
                downloaded_bytes += len(chunk)

        # Validation: file not empty
        if downloaded_bytes == 0:
            os.remove(temp_path)
            raise Exception("La descarga resultó en un archivo vacío.")

        return temp_path

    def _create_handover_bat(self, temp_path: str) -> str:
        """
        Generates the .bat that:
        1. Waits for the current process to end.
        2. Replaces the .exe.
        3. Launches the new version.
        4. Self-deletes.
        """
        bat_path = os.path.join(self.directory, "update.bat")

        # Escape paths to avoid issues with special characters
        exe_path = self.current_exe.replace('"', '""')
        temp_file = temp_path.replace('"', '""')

        content = (
            '@echo off\n'
            'timeout /t 2 /nobreak > nul\n'
            f'del /f /q "{exe_path}"\n'
            f'move /y "{temp_file}" "{exe_path}"\n'
            f'start "" "{exe_path}"\n'
            'del "%~f0"\n'
        )

        with open(bat_path, "w", encoding="utf-8") as f:
            f.write(content)

        return bat_path