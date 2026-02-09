# pptx-slide-copier

PowerPoint slide copier utility - Deep copy slides between presentations while preserving all formatting, images, and layouts.

## Features

- Deep copy slides from one PowerPoint presentation to another
- Preserves all formatting, fonts, colors, and styles
- Copies images with correct relationship IDs
- Matches slide layouts via slide masters (supports custom layouts)
- Compatible with python-pptx 1.0.2+ API
- Handles themes and layout preservation

## Installation

```bash
pip install -e .
```

For development dependencies:

```bash
pip install -e ".[dev]"
```

## Quick Start

```python
from pptx import Presentation
from pptx_slide_copier import SlideCopier

# Open source and target presentations
source_prs = Presentation("source.pptx")
target_prs = Presentation("target.pptx")

# Copy slide 0 from source to target
copied_slide = SlideCopier.copy_slide(source_prs, 0, target_prs)

# Save the target presentation
target_prs.save("output.pptx")
```

## Usage

### Basic Slide Copying

```python
from pptx import Presentation
from pptx_slide_copier import SlideCopier

# Load presentations
source = Presentation("template.pptx")
target = Presentation()  # Create new presentation

# Copy first slide (index 0)
SlideCopier.copy_slide(source, 0, target)

# Copy multiple slides
for i in range(3):
    SlideCopier.copy_slide(source, i, target)

target.save("output.pptx")
```

### Copying to Presentation Based on Template

For best results (to preserve themes and layouts), create the target presentation from the same template:

```python
from pptx import Presentation
from pptx_slide_copier import SlideCopier

# Load template
template = Presentation("template.pptx")

# Create target from same template to preserve theme
target = Presentation("template.pptx")

# Remove all slides from target
while len(target.slides) > 0:
    rId = target.slides._sldIdLst[0].rId
    target.part.drop_rel(rId)
    del target.slides._sldIdLst[0]

# Now copy slides
SlideCopier.copy_slide(template, 0, target)

target.save("output.pptx")
```

## How It Works

### Deep Copying at XML Level

The library uses `deepcopy` at the XML element level to ensure all shape properties, formatting, and attributes are preserved:

```python
from copy import deepcopy
new_element = deepcopy(shape.element)
dest_slide.shapes._spTree.insert_element_before(new_element, "p:extLst")
```

### Image Relationship Mapping

Images are copied with proper relationship ID mapping to ensure they display correctly:

1. Copy image data from source slide
2. Add image to target presentation
3. Map old relationship IDs to new ones
4. Update XML references in copied shapes

### Layout Matching

Slide layouts are matched via slide masters to support custom layouts:

1. Find source slide's layout and master
2. Match master by name in target presentation
3. Find matching layout within the matched master
4. Fall back to index-based matching if name match fails

## Technical Details

- **python-pptx 1.0.2+ compatibility**: Handles both old and new API for `get_or_add_image_part()`
- **XML namespaces**: Correctly handles presentationml, drawingml, and relationship namespaces
- **Theme preservation**: Works best when target presentation is created from same template
- **Error handling**: Gracefully continues if individual shapes or images fail to copy

## Testing

```bash
pytest
```

With coverage:

```bash
pytest --cov=pptx_slide_copier --cov-report=html
```

## License

MIT

## Author

SpringMT
