# pptx2png

A one-click Python library for converting `.pptx` files into `.png` images.

**Installation**:

```cmd
pip install pptx2png

```

**Example Code**:

```python
import pptx2png

# pptx: required, the PowerPoint file to process
# output_dir: optional, default is ./pptx2png in the same directory
# slide_range: optional, specifies the range of slides to convert
# scale: optional, resolution scale.
#        If not specified, it defaults to screen resolution.
pptx2png.topng(
    pptx="your_presentation.pptx",
    output_dir="./output",
    slide_range=[1, 5],
    scale=2
)

pptx2png.whatis() # print info
```

> A graphical EXE version is also available. See [GitHub Releases](https://github.com/Water-Run/pptx2png/releases/tag/pptx2png) for more information.
