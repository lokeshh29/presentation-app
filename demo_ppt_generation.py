"""
Demo script for PPTGenerator - PowerPoint Generation Basics

This script demonstrates how to use all the basic functions of the PPTGenerator class:
- add_slide()
- delete_slide(n)
- modify_layout()
- insert_chart()
- insert_image()
- update_text()
- change_background()

Run this script to create a sample presentation with various elements.
"""

from ppt_generator import PPTGenerator
import os


def create_sample_image():
    """Create a simple sample image for demonstration (optional)."""
    try:
        from PIL import Image, ImageDraw, ImageFont
        
        # Create a simple sample image
        img = Image.new('RGB', (400, 300), color='lightblue')
        draw = ImageDraw.Draw(img)
        
        # Add some text to the image
        try:
            # Try to use default font
            font = ImageFont.load_default()
        except:
            font = None
            
        draw.text((150, 130), "Sample Image", fill='darkblue', font=font)
        draw.rectangle([50, 50, 350, 250], outline='navy', width=3)
        
        img.save('sample_image.png')
        print("Created sample image: sample_image.png")
        return True
        
    except ImportError:
        print("PIL/Pillow not installed. Skipping image creation.")
        print("Install with: pip install Pillow")
        return False
    except Exception as e:
        print(f"Could not create sample image: {str(e)}")
        return False


def demo_ppt_generation():
    """Demonstrate all PPT generation functions."""
    
    print("=== PowerPoint Generation Demo ===\n")
    
    # Initialize PPT Generator
    ppt = PPTGenerator()
    print("✓ PPTGenerator initialized\n")
    
    # Show available layouts
    print("Available slide layouts:")
    for layout in ppt.get_available_layouts():
        print(f"  {layout}")
    print()
    
    # 1. Add slides with different layouts
    print("1. Adding slides with different layouts...")
    
    # Title slide
    slide0 = ppt.add_slide(layout_index=0, title="PowerPoint Generation Demo", 
                          subtitle="Created with python-pptx")
    
    # Content slide
    slide1 = ppt.add_slide(layout_index=1, title="Features Overview", 
                          subtitle="• Dynamic slide creation\n• Chart insertion\n• Image handling\n• Text formatting")
    
    # Two content slide
    slide2 = ppt.add_slide(layout_index=3, title="Comparison Slide")
    
    # Blank slide
    slide3 = ppt.add_slide(layout_index=6)
    
    print(f"✓ Added {ppt.get_slide_count()} slides\n")
    
    # 2. Update text with formatting
    print("2. Updating text content with formatting...")
    ppt.update_text(slide_index=1, 
                   text_updates={
                       'title': 'Enhanced Features Overview',
                       'content': '• Professional slide layouts\n• Interactive charts and graphs\n• High-quality image integration\n• Advanced text formatting'
                   },
                   font_size=14,
                   font_color=(50, 50, 150),
                   bold=True,
                   alignment='left')
    print("✓ Updated text with custom formatting\n")
    
    # 3. Insert chart
    print("3. Inserting charts...")
    chart_data = {
        'categories': ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024'],
        'series': [
            {'name': 'Revenue', 'values': [100, 125, 110, 140]},
            {'name': 'Profit', 'values': [20, 30, 25, 35]}
        ]
    }
    
    ppt.insert_chart(slide_index=2, chart_type='column', data=chart_data, 
                    position=(1, 3), size=(6, 4))
    print("✓ Added column chart to slide 3\n")
    
    # 4. Insert image (if available)
    print("4. Inserting images...")
    image_created = create_sample_image()
    if image_created and os.path.exists('sample_image.png'):
        ppt.insert_image(slide_index=3, image_path='sample_image.png', 
                        position=(2, 2), size=(6, 4))
        print("✓ Added image to slide 4\n")
    else:
        print("! Skipped image insertion (no sample image available)\n")
    
    # 5. Change backgrounds
    print("5. Changing slide backgrounds...")
    
    # Solid color background for slide 1
    ppt.change_background(slide_index=0, background_type="solid", 
                         color=(240, 248, 255))  # Light blue
    print("✓ Applied light blue background to title slide")
    
    # Image background for slide 4 (if image exists)
    if image_created and os.path.exists('sample_image.png'):
        ppt.change_background(slide_index=3, background_type="image", 
                             image_path='sample_image.png')
        print("✓ Applied image background to slide 4")
    
    print()
    
    # 6. Modify layout
    print("6. Modifying slide layouts...")
    ppt.modify_layout(slide_index=2, new_layout_index=1)  # Change to Title and Content
    print("✓ Changed slide 3 layout to 'Title and Content'\n")
    
    # 7. Demonstrate delete slide (add extra slide first)
    print("7. Demonstrating slide deletion...")
    extra_slide = ppt.add_slide(layout_index=6, title="Temporary Slide")
    print(f"Added temporary slide. Total slides: {ppt.get_slide_count()}")
    
    ppt.delete_slide(ppt.get_slide_count() - 1)  # Delete the last slide
    print(f"Deleted temporary slide. Total slides: {ppt.get_slide_count()}\n")
    
    # Save the presentation
    print("8. Saving presentation...")
    ppt.save('demo_presentation.pptx')
    print("✓ Presentation saved as 'demo_presentation.pptx'\n")
    
    # Summary
    print("=== Demo Complete ===")
    print("Functions demonstrated:")
    print("✓ add_slide() - Added slides with different layouts")
    print("✓ update_text() - Updated text with formatting")
    print("✓ insert_chart() - Added column chart")
    print("✓ insert_image() - Added image (if available)")
    print("✓ change_background() - Applied solid and image backgrounds")
    print("✓ modify_layout() - Changed slide layout")
    print("✓ delete_slide() - Removed slide")
    print("\nOpen 'demo_presentation.pptx' to view the results!")


if __name__ == "__main__":
    try:
        demo_ppt_generation()
    except Exception as e:
        print(f"Error running demo: {str(e)}")
        print("Make sure python-pptx is installed: pip install python-pptx")