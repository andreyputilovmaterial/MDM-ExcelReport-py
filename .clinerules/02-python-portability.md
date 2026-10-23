# Python code style

It must be usable as an imported module *or* included as a git submodule, bundled via
`pinliner`, so it obeys strict portability constraints:

1. Relative imports only — entry-point scripts that are also run directly are the exception (see below),
2. Liftable to a different location in a bigger project without changes,
3. Every folder on the path to every module contains `__init__.py`,
4. No static files/assets used as-is at runtime — static assets are discouraged and,
   if needed, must be embedded into a generated Python module at build time (see below),
5. Minimal dependencies.

## Entry-point scripts and conditional imports

This project can be used in different ways: distributed as a single-file bundle created
with `pinliner`, or called directly during development. To support both, a few scripts
are also launched as a program in development and therefore use a three-way conditional
import block instead of a single import style:

```python
if __name__ == '__main__':
    # run as a program, e.g. `python src/report_create.py`
    from sibling_module import some_function
elif '.' in __name__:
    # imported as a package member
    from .sibling_module import some_function
else:
    # included in the pinliner bundle with no parent package
    from sibling_module import some_function
```

The module and function names above are illustrative — replace them with the real
sibling module of your entry point. In this repo the scripts that follow this pattern
are `src/launcher.py` and `src/report_create.py`. Any script that is meant to be
launched directly should use this pattern; every other module (the ones that are only
ever imported) uses plain relative imports.

## Static assets and the `GENERATED` folder

The default distributable is a single-file bundle produced with `pinliner`, which can
only embed Python modules. A plain static file (a version string, an HTML/CSS/JS
template, or anything that is not a python module) cannot be used as-is in that bundle,
so:

1. **Static assets are discouraged.** It is easier to avoid them than to solve their
   embedding into the pinliner bundle.
2. **When a static asset is necessary — or clearly beneficial — it is allowed, but only
   via this technique:** the build script converts the asset into a Python module
   written to a `GENERATED` subfolder, which the real source code then imports like any
   other module.

This repo currently has **no static assets**: it produces Excel workbooks from plain
Python, so there is nothing to embed. The technique above is kept here for consistency
with the sibling packages (e.g. `mdmreportpy` embeds its HTML/CSS/JS templates this
way, and the other packages embed their git version string via `src/GENERATED/_VERSION.py`).

## When inspecting source code

Any code under a `GENERATED` folder is **generated output, not maintained source code**:

- It exists only to solve the bundling/embedding problem described above.
- It is produced by the build script and can be regenerated at any time.
- It is **not tracked in git** (`.gitignore` excludes `**/GENERATED/`).
- Do not edit it by hand; change the original asset or the build script that produces it
  instead.

