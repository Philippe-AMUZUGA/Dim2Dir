"""Dim2Dir : projection et visualisation de plans d'expériences."""
from .core import (
    Dim2DirError, Projection, Result, Style, Limits, __version__,
    project, figure, visualize, project_fractal_2d,
    add_visualization_to_sheet, run_visualization, build_output_path, main,
)

__all__ = [
    "Dim2DirError", "Projection", "Result", "Style", "Limits", "__version__",
    "project", "figure", "visualize", "project_fractal_2d",
    "add_visualization_to_sheet", "run_visualization", "build_output_path", "main",
]
