import os
import win32com.client
import gc

def get_interactive_inputs():
    """Prompts the user interactively in the console for input path, output path, and DPI."""
    print("=" * 60)
    print(" POWERPOINT TO PNG BULK CONVERTER (Interactive Mode)")
    print("=" * 60)

    # 1. Input directory prompt
    default_input = r"D:\Powerpoint\Input"
    raw_input = input(f"Enter PowerPoint input directory [Default: {default_input}]: ").strip().strip('"')
    input_directory = raw_input if raw_input else default_input

    # 2. Output directory prompt
    default_output = r"D:\Powerpoint\Output"
    raw_output = input(f"Enter PNG output directory [Default: {default_output}]: ").strip().strip('"')
    output_directory = raw_output if raw_output else default_output

    # 3. DPI prompt
    default_dpi = 300
    raw_dpi = input(f"Enter target DPI resolution [Default: {default_dpi}]: ").strip()
    try:
        dpi = int(raw_dpi) if raw_dpi else default_dpi
    except ValueError:
        print(f"Invalid integer entered for DPI. Using default value: {default_dpi}")
        dpi = default_dpi

    print("-" * 60)
    return input_directory, output_directory, dpi

def convert_ppts_to_pngs(input_dir, output_dir, target_dpi):
    if not os.path.exists(input_dir):
        print(f"Error: Input directory not found at {input_dir}")
        return

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created master output directory: {output_dir}")

    # Find all PowerPoint files in the directory (ignore temporary ~$ files)
    ppt_files = [f for f in os.listdir(input_dir) if f.lower().endswith(('.ppt', '.pptx')) and not f.startswith('~$')]

    if not ppt_files:
        print(f"No PowerPoint files found in {input_dir}")
        return

    print(f"Found {len(ppt_files)} presentations. Booting PowerPoint...")

    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        
        for filename in ppt_files:
            input_ppt = os.path.join(input_dir, filename)
            
            # Create a dedicated subfolder for this presentation's slides
            base_name = os.path.splitext(filename)[0]
            pres_output_folder = os.path.join(output_dir, base_name)
            
            if not os.path.exists(pres_output_folder):
                os.makedirs(pres_output_folder)

            print(f"\nProcessing: {filename}")
            
            try:
                # Open presentation in background
                presentation = powerpoint.Presentations.Open(input_ppt, WithWindow=False)

                # PowerPoint dimensions are measured in points (72 points = 1 inch)
                # We calculate the required pixel resolution based on your target DPI
                slide_width_pt = presentation.PageSetup.SlideWidth
                slide_height_pt = presentation.PageSetup.SlideHeight

                width_px = int((slide_width_pt / 72) * target_dpi)
                height_px = int((slide_height_pt / 72) * target_dpi)

                print(f"  Target Resolution: {width_px}x{height_px} pixels ({target_dpi} DPI)")

                # Export each slide directly to PNG
                for i, slide in enumerate(presentation.Slides):
                    slide_num = i + 1
                    output_path = os.path.join(pres_output_folder, f"Slide_{slide_num:02d}.png")
                    
                    # Direct export overrides PowerPoint's default low-res settings
                    slide.Export(output_path, "PNG", width_px, height_px)
                    print(f"  Saved: Slide_{slide_num:02d}.png")

                presentation.Close()

            except Exception as e:
                print(f"  Failed to process {filename}: {e}")

    except Exception as e:
        print(f"Failed to start PowerPoint application: {e}")
        
    finally:
        # Guarantee PowerPoint closes to prevent background kernel hangs
        if 'powerpoint' in locals() and powerpoint:
            powerpoint.Quit()
        gc.collect()
        print("\nSUCCESS! All presentations processed.")

if __name__ == "__main__":
    input_dir, output_dir, target_dpi = get_interactive_inputs()
    convert_ppts_to_pngs(input_dir, output_dir, target_dpi)