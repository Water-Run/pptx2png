import os
import sys
import ctypes

# Try to import win32com, prompt user if missing
try:
    import win32com.client
except ImportError:
    print("Error: Library 'pywin32' is required. Please install it via: pip install pywin32")
    sys.exit(1)

def topng(pptx, output_dir="./output", slide_range=None, scale=None):
    """
    Convert PowerPoint slides to PNG images.

    Args:
        pptx (str): Path to the .pptx file.
        output_dir (str): Directory to save the images. Default is './output'.
        slide_range (list): Optional. A list [start, end] specifying slide range (1-based).
                            Example: [1, 5] converts slides 1 to 5.
        scale (int): Optional. Resolution scale.
                     If None or 0, it adapts to the screen's long edge resolution.
                     If specified (e.g., 1, 2), it scales relative to original slide points.
    """
    # 1. Path handling
    pptx_path = os.path.abspath(pptx)
    output_path = os.path.abspath(output_dir)

    # Safety Check: Prevent overwriting source code directory if names clash
    # Get the directory where this script resides
    current_script_dir = os.path.dirname(os.path.abspath(__file__))
    if output_path == current_script_dir:
        print("Error: Output directory cannot be the same as the library source directory.")
        return

    if not os.path.exists(pptx_path):
        print("Error: File '%s' not found." % pptx_path)
        return

    if not os.path.exists(output_path):
        try:
            os.makedirs(output_path)
            print("Created output directory: %s" % output_path)
        except OSError as e:
            print("Error: Could not create output directory. %s" % e)
            return

    # 2. Initialize PowerPoint Application
    powerpoint = None
    presentation = None
    try:
        # Use DispatchEx to ensure a fresh instance if needed, or Dispatch for shared
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    except Exception as e:
        print("Error: Could not initialize PowerPoint. Make sure Microsoft PowerPoint is installed.")
        print("Details: %s" % e)
        return

    try:
        # 3. Open Presentation (WithWindow=False attempts background processing)
        # Note: Some PPT versions force visibility despite this flag.
        presentation = powerpoint.Presentations.Open(pptx_path, WithWindow=False)

        # 4. Determine Slide Range
        total_slides = presentation.Slides.Count
        start_slide = 1
        end_slide = total_slides

        # Validate and apply slide_range if provided
        if slide_range and isinstance(slide_range, list) and len(slide_range) == 2:
            # Ensure start is at least 1
            s_req = max(1, slide_range[0])
            # Ensure end is at most total_slides
            e_req = min(total_slides, slide_range[1])
            
            if s_req <= e_req:
                start_slide = s_req
                end_slide = e_req

        # 5. Calculate Target Resolution
        slide_width = presentation.PageSetup.SlideWidth
        slide_height = presentation.PageSetup.SlideHeight
        
        target_w = 0
        target_h = 0

        # Logic: If scale is not provided, use screen resolution (Long Edge) with a boost
        if not scale:
            try:
                user32 = ctypes.windll.user32
                screen_w = user32.GetSystemMetrics(0)
                screen_h = user32.GetSystemMetrics(1)

                # Long edge of the screen
                screen_long = max(screen_w, screen_h)

                boost = float(os.getenv("PPTX2PNG_SCREEN_SCALE", 2))
                if boost <= 0:
                    boost = 2

                target_long = int(screen_long * boost)

                # Calculate aspect ratio of the slide
                slide_ratio = slide_width / slide_height

                if slide_ratio >= 1:  # Landscape Slide
                    target_w = target_long
                    target_h = int(target_long / slide_ratio)
                else:  # Portrait Slide
                    target_h = target_long
                    target_w = int(target_long * slide_ratio)

                print(
                    f"Mode: Auto-Resolution (Screen long edge {screen_long}px, "
                    f"boost {boost}x -> target long {target_long}px)"
                )
            except Exception:
                # Fallback if ctypes fails
                target_w = int(slide_width * 2)
                target_h = int(slide_height * 2)
                print("Mode: Fallback Resolution (2x)")
        else:
            # Manual scale
            target_w = int(slide_width * scale)
            target_h = int(slide_height * scale)
            print("Mode: Manual Scale (%dx)" % scale)

        print("Processing '%s'..." % os.path.basename(pptx))
        print("Target Size: %dx%d px" % (target_w, target_h))
        print("Converting slides %d to %d..." % (start_slide, end_slide))

        # 6. Iterate and Export
        count = 0
        # CRITICAL FIX: 'range' here now refers to the built-in function, 
        # because the argument was renamed to 'slide_range'
        for i in range(start_slide, end_slide + 1):
            slide = presentation.Slides(i)
            # Filename format: Slide_1.png, Slide_2.png
            image_name = "Slide_%d.png" % i
            image_path = os.path.join(output_path, image_name)

            # Export to PNG
            slide.Export(image_path, "PNG", target_w, target_h)
            count += 1
            print("Saved: %s" % image_name)

        print("Done! %d images saved to '%s'." % (count, output_path))

    except Exception as e:
        print("An error occurred during conversion: %s" % e)
        # Import traceback to print full stack trace for debugging
        import traceback
        traceback.print_exc()
        
    finally:
        # 7. Cleanup Resources
        if presentation:
            try:
                presentation.Close()
            except Exception:
                pass
        # powerpoint.Quit() is intentionally omitted to avoid closing user's active windows
        pass

def whatis():
    """Prints the library information."""
    info = """
--------------------------------------------------
pptx2png Info
--------------------------------------------------
Version     : 2025.1
Author      : WaterRun
GitHub      : https://github.com/Water-Run/pptx2png
Email       : 2263633954@qq.com
Description : A library to convert PPTX to PNG.

Note: A Windows GUI (EXE) version is also available:
https://github.com/Water-Run/pptx2png/releases/tag/pptx2png
--------------------------------------------------
    """
    print(info.strip())
    