# OOXML Slide Copy: Internal Mechanics and Bug Fixes

This document records the OOXML specification details, python-pptx internal behavior, and the bugs found and fixed in `pptx-slide-copier`.

## OOXML Presentation Structure Overview

A `.pptx` file is a ZIP archive containing XML files following the Office Open XML (OOXML) specification.  The key structural files relevant to slide copying are:

```
pptx (ZIP)
├── [Content_Types].xml          # Maps file extensions/paths to content types
├── ppt/
│   ├── presentation.xml         # Main presentation definition
│   ├── _rels/presentation.xml.rels  # Relationships for presentation.xml
│   ├── slides/
│   │   ├── slide1.xml           # Slide content
│   │   └── _rels/slide1.xml.rels   # Slide relationships (layout, images, etc.)
│   ├── slideLayouts/
│   │   ├── slideLayout1.xml     # Layout definitions
│   │   └── _rels/slideLayout1.xml.rels
│   ├── slideMasters/
│   │   ├── slideMaster1.xml     # Master slide definitions
│   │   └── _rels/slideMaster1.xml.rels
│   ├── theme/
│   │   └── theme1.xml           # Theme definitions
│   └── media/
│       └── image1.png           # Embedded images
```

### Slide Ordering: `sldIdLst`

Slide order is determined by the `<p:sldIdLst>` element in `presentation.xml`:

```xml
<p:sldIdLst>
  <p:sldId id="256" r:id="rId7"/>   <!-- 1st slide -->
  <p:sldId id="257" r:id="rId8"/>   <!-- 2nd slide -->
  <p:sldId id="258" r:id="rId9"/>   <!-- 3rd slide -->
</p:sldIdLst>
```

- `id`: A unique numeric identifier for the slide (range: 256 to 2,147,483,647). This is **persistent** and does **not** change when slides are reordered.
- `r:id`: A relationship ID pointing to the actual slide XML file via `presentation.xml.rels`.

The position of `<p:sldId>` elements within `<p:sldIdLst>` determines the visible slide order. Reordering these elements is the correct way to change slide order.

### Master and Layout ID Space: `sldMasterIdLst` and `sldLayoutIdLst`

```xml
<!-- In presentation.xml -->
<p:sldMasterIdLst>
  <p:sldMasterId id="2147483648" r:id="rId1"/>  <!-- id is REQUIRED -->
</p:sldMasterIdLst>

<!-- In slideMaster1.xml -->
<p:sldLayoutIdLst>
  <p:sldLayoutId id="2147483649" r:id="rId1"/>  <!-- id is REQUIRED -->
  <p:sldLayoutId id="2147483650" r:id="rId2"/>
</p:sldLayoutIdLst>
```

- Both `sldMasterId` and `sldLayoutId` elements share the **same id space**.
- Valid range: `2147483648` (`0x80000000`) and above.
- The `id` attribute is **required** by the OOXML specification. PowerPoint will report the file as corrupt if it is missing.
- IDs must be **unique** across all `sldMasterId` and `sldLayoutId` elements in the entire presentation.

### Relationship Chain

Each slide connects to the rest of the structure via relationships:

```
slide1.xml  --[slideLayout]--> slideLayout1.xml
                                  --[slideMaster]--> slideMaster1.xml
                                                        --[theme]--> theme1.xml
                                                        --[image]--> media/image1.png
```

## python-pptx Internal Behavior

### `add_slide()` Flow

```python
Slides.add_slide(slide_layout)
  └── PresentationPart.add_slide(slide_layout)
      ├── Creates new SlidePart with unique partname
      ├── Creates relationship: PresentationPart → SlidePart
      └── Returns (rId, slide)
  └── CT_SlideIdList.add_sldId(rId)
      └── Creates <p:sldId id=next_id rId=rId/>
```

`add_slide()` always **appends** the new slide to the end of `sldIdLst`.

### Part Visibility via `iter_parts()`

`Package.iter_parts()` traverses the relationship graph starting from the root part. A part is only visible if it is **reachable** through relationships from the root.

This means:
- A newly created `SlideMasterPart` is **not visible** in `iter_parts()` until it is connected to the presentation via `relate_to()`.
- `Package.next_image_partname()` and `get_or_add_image_part()` rely on `iter_parts()` for deduplication.
- If images are added to an unregistered part, deduplication fails and **duplicate partnames** are generated.

### `CT_SlideMasterIdListEntry` Limitation

python-pptx's `CT_SlideMasterIdListEntry` class only defines `rId` as a `RequiredAttribute`, not `id`:

```python
class CT_SlideMasterIdListEntry(BaseOxmlElement):
    rId: str = RequiredAttribute("r:id", XsdString)
    # Note: 'id' is NOT declared as an attribute
```

This means `_add_sldMasterId(rId=rId)` creates an element **without** the `id` attribute. The `id` must be set manually via `element.set("id", value)`.

## Bugs Found and Fixed

### Bug 1: Missing `id` Attribute on `sldMasterId` and `sldLayoutId`

**Symptom**: PowerPoint reports the file as corrupt and requires repair.

**Root cause**: When copying a slide master or layout, the code called `_add_sldMasterId(rId=rId)` and `_add_sldLayoutId(rId=rId)` without setting the `id` attribute. Since python-pptx does not define `id` as a managed attribute on these element classes, the resulting XML was:

```xml
<!-- INVALID: missing id attribute -->
<p:sldMasterId r:id="rId9"/>
<p:sldLayoutId r:id="rId2"/>
```

**Fix**: After calling `_add_sldMasterId`/`_add_sldLayoutId`, explicitly set the `id` attribute using a helper method `_next_unique_id()` that scans all existing IDs and returns the next available value:

```python
sld_master_id = sld_master_id_lst._add_sldMasterId(rId=rId)
sld_master_id.set("id", str(SlideCopier._next_unique_id(target_prs)))
```

**Affected methods**: `_copy_slide_master_part`, `_copy_slide_layout_part`

### Bug 2: Duplicate Image Parts (ZIP Corruption)

**Symptom**: ZIP file contains duplicate entries (e.g., `ppt/media/image15.png` appears twice). Some PPTX readers fail to open the file; others silently use one copy and lose the other image.

**Root cause**: In `_copy_slide_master_part` and `_copy_slide_layout_part`, the part registration step (connecting the new part to the presentation via `relate_to`) happened **after** the image copying step (`_copy_part_rels`). Since the new part was not yet reachable via `iter_parts()`, `get_or_add_image_part()` could not find previously copied images for deduplication. This caused `ImagePart.new()` to generate the same partname for different images.

**Timeline of the bug**:

```
1. Create new SlideMasterPart (not yet registered in package)
2. Copy image A → get_or_add_image_part → SHA1 not found → ImagePart.new("image15.png")
   (image15.png is added to master's rels, but master is not in package graph)
3. Copy image B → get_or_add_image_part → SHA1 not found (can't see image15.png)
   → ImagePart.new("image15.png") again!
4. Register master in presentation (too late)
```

**Fix**: Moved the part registration step (`relate_to` + `sldMasterIdLst`/`sldLayoutIdLst` update) to **before** the image copying step:

```
1. Create new SlideMasterPart
2. Register master in presentation (relate_to) ← moved earlier
3. Copy image A → get_or_add_image_part → creates image15.png (now visible)
4. Copy image B → get_or_add_image_part → creates image16.png (correct!)
```

**Affected methods**: `_copy_slide_master_part`, `_copy_slide_layout_part`

### Bug 3: Slide Size Overwritten

**Symptom**: All existing slides in the target presentation display at the wrong aspect ratio or size after copying.

**Root cause**: `_copy_slide_size()` was called unconditionally in `copy_slide()`, overwriting the target's slide dimensions with the source's. When the source uses a different slide size (e.g., 16:9 vs 4:3), all existing target slides are affected.

```python
# BEFORE (always overwrites):
SlideCopier._copy_slide_size(source_prs, target_prs)
```

**Fix**: Only copy slide size when the target has no existing slides:

```python
if len(target_prs.slides) == 0:
    SlideCopier._copy_slide_size(source_prs, target_prs)
```

**Affected method**: `copy_slide`

### Bug 4: On-Demand Path Creates Unnecessary Duplicate Masters

**Symptom**: When using `copy_slide()` directly (without `copy_slides()`), a duplicate slide master is created even when the source and target share the same theme. This bloats the file and can confuse PowerPoint.

**Root cause**: `_get_or_copy_slide_master()` did not check whether the target already had a master with the same theme. The `_find_matching_master()` method existed but was only used in `copy_layouts()`. The on-demand path always created a new master.

**Fix**: Added `_find_matching_master()` check to `_get_or_copy_slide_master()`:

```python
matching_master = SlideCopier._find_matching_master(source_master_part, target_prs)
if matching_master is not None:
    target_master_part = matching_master.part
    cache[cache_key] = target_master_part
    return target_master_part
```

**Affected method**: `_get_or_copy_slide_master`

### Bug 5: Theme Background Images Lost

**Symptom**: Copied slides lose their theme-defined background. The background appears as a solid color or is blank instead of showing the expected image.

**Root cause**: `_copy_theme_part()` copied the theme XML as a blob but did not copy the theme's `.rels` file. OOXML themes can contain `r:embed` references to images (e.g., for background fills defined in `<a:bgFillStyleLst>`). These references are resolved through the theme's `.rels` file:

```xml
<!-- In theme1.xml -->
<a:bgFillStyleLst>
  <a:blipFill rotWithShape="1">
    <a:blip r:embed="rId1"/>   <!-- references theme1.xml.rels -->
    <a:stretch/>
  </a:blipFill>
</a:bgFillStyleLst>

<!-- In theme1.xml.rels -->
<Relationship Id="rId1" Type="...image" Target="../media/image1.jpeg"/>
```

When the theme was copied without its `.rels`, the `r:embed="rId1"` reference became dangling, and PowerPoint could not find the background image.

**Fix**: Added relationship copying logic to `_copy_theme_part()`. Since theme parts are plain blob-based `Part` objects (not `BaseSlidePart`), the standard `_copy_part_rels` cannot be used. Instead, each relationship is copied individually: image blobs are duplicated as new `Part` objects, and if any rIds change during copy, the theme XML blob is patched with the new rId values.

**Affected method**: `_copy_theme_part`

## Lessons Learned

1. **Always validate with real PowerPoint files.** python-pptx can read files that PowerPoint rejects. Unit tests with `Presentation()` objects miss issues that only surface with production files.

2. **Part registration order matters.** In python-pptx, `Package.iter_parts()` walks the relationship graph. Parts must be connected to the graph before any operation that depends on `iter_parts()` (such as image deduplication).

3. **python-pptx's OxmlElement classes may not cover all required XML attributes.** Always verify the generated XML against the OOXML specification, especially for `id` attributes.

4. **Slide size is a presentation-level property.** Overwriting it affects all existing slides, not just the copied one.

5. **Theme parts have their own relationships.** Themes are not self-contained XML blobs. They can reference external resources (images, fonts) via `.rels` files. When copying a theme, its relationships must be copied too.
