# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information

project = 'Excel-to-JSON'
copyright = '2022~2025, WTSolutions'
author = 'WTSolutions'
release = '3.1.0.0'

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration


extensions = ['myst_parser','sphinx_sitemap']
html_baseurl = 'https://excel-to-json.wtsolutions.cn/zh-cn/latest/'
sitemap_url_scheme = "{link}"
html_extra_path = ['robots.txt','ads.txt']
html_js_files = ['custom.js']

templates_path = ['_templates']
exclude_patterns = [
    'donation.md','usage.md','requirements.txt','.DS_Store'
]
language = 'zh-cn'
source_suffix = {
    '.rst': 'restructuredtext',
    '.txt': 'markdown',
    '.md': 'markdown',
}

# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

# html_theme = 'alabaster'
html_theme = 'sphinx_rtd_theme'
html_static_path = ['_static']
