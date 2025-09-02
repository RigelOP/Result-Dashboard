"""Adapter module: re-export the extractor functions from Result_Downloader.

This keeps `dashboard_app.py` imports tidy while using the existing
`Result_Downloader.py` implementation.
"""

try:
	from Result_Downloader import process_single_student, save_results_to_excel
except Exception as e:
	# Provide clear import-time error for easier debugging
	raise ImportError(f"Unable to import extractor functions from Result_Downloader: {e}")

__all__ = ["process_single_student", "save_results_to_excel"]

