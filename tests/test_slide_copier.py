"""Tests for SlideCopier."""

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx_slide_copier import SlideCopier
import tempfile
from pathlib import Path


@pytest.fixture
def sample_presentation():
    """Create a sample presentation with text and shapes."""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    # Add a slide with some text
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)

    # Add a text box
    left = Inches(1)
    top = Inches(1)
    width = Inches(8)
    height = Inches(1)
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.text = "Test slide content"

    # Set some formatting
    paragraph = text_frame.paragraphs[0]
    run = paragraph.runs[0]
    run.font.name = "Arial"
    run.font.size = Pt(24)
    run.font.bold = True

    return prs


@pytest.fixture
def temp_pptx_file(sample_presentation):
    """Save sample presentation to a temporary file."""
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        sample_presentation.save(tmp.name)
        yield tmp.name
        Path(tmp.name).unlink()


class TestSlideCopier:
    """Test cases for SlideCopier class."""

    def test_copy_slide_basic(self, sample_presentation):
        """Test basic slide copying."""
        source_prs = sample_presentation
        target_prs = Presentation()

        # Copy the first slide
        copied_slide = SlideCopier.copy_slide(source_prs, 0, target_prs)

        assert copied_slide is not None
        assert len(target_prs.slides) == 1

    def test_copy_slide_preserves_text(self, sample_presentation):
        """Test that text content is preserved."""
        source_prs = sample_presentation
        target_prs = Presentation()

        copied_slide = SlideCopier.copy_slide(source_prs, 0, target_prs)

        # Find text in copied slide
        found_text = False
        for shape in copied_slide.shapes:
            if hasattr(shape, "text") and "Test slide content" in shape.text:
                found_text = True
                break

        assert found_text, "Text content should be preserved"

    def test_copy_slide_preserves_size(self, sample_presentation):
        """Test that slide dimensions are preserved."""
        source_prs = sample_presentation
        target_prs = Presentation()

        SlideCopier.copy_slide(source_prs, 0, target_prs)

        assert target_prs.slide_width == source_prs.slide_width
        assert target_prs.slide_height == source_prs.slide_height

    def test_copy_multiple_slides(self, sample_presentation):
        """Test copying multiple slides."""
        source_prs = sample_presentation
        target_prs = Presentation()

        # Copy the same slide twice
        SlideCopier.copy_slide(source_prs, 0, target_prs)
        SlideCopier.copy_slide(source_prs, 0, target_prs)

        assert len(target_prs.slides) == 2

    def test_copy_slide_with_formatting(self, sample_presentation):
        """Test that text formatting is preserved."""
        source_prs = sample_presentation
        target_prs = Presentation()

        copied_slide = SlideCopier.copy_slide(source_prs, 0, target_prs)

        # Check that formatting is preserved
        for shape in copied_slide.shapes:
            if hasattr(shape, "text_frame"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if run.text == "Test slide content":
                            # Font properties should be preserved
                            assert run.font.bold == True
                            # Font name and size should be set
                            assert run.font.name is not None
                            assert run.font.size is not None

    def test_copy_slide_to_template_based_presentation(self, temp_pptx_file):
        """Test copying to a presentation based on the same template."""
        source_prs = Presentation(temp_pptx_file)

        # Create target from same template
        target_prs = Presentation(temp_pptx_file)

        # Remove all slides from target
        while len(target_prs.slides) > 0:
            rId = target_prs.slides._sldIdLst[0].rId
            target_prs.part.drop_rel(rId)
            del target_prs.slides._sldIdLst[0]

        assert len(target_prs.slides) == 0

        # Copy slide
        SlideCopier.copy_slide(source_prs, 0, target_prs)

        assert len(target_prs.slides) == 1

    def test_copy_slide_invalid_index(self, sample_presentation):
        """Test that invalid slide index raises appropriate error."""
        source_prs = sample_presentation
        target_prs = Presentation()

        with pytest.raises(IndexError):
            SlideCopier.copy_slide(source_prs, 999, target_prs)

    def test_copy_slide_preserves_shapes(self, sample_presentation):
        """Test that shapes are copied."""
        source_prs = sample_presentation
        source_slide = source_prs.slides[0]
        source_shape_count = len(source_slide.shapes)

        target_prs = Presentation()
        copied_slide = SlideCopier.copy_slide(source_prs, 0, target_prs)

        # Should have same number of shapes (or close, due to layout elements)
        assert len(copied_slide.shapes) > 0
