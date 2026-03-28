# CropImage

`CropImage.ps1` crops a region from an image file and saves the result as a JPEG file.

The script supports:

- cropping by `x`, `y`, `width`, and `height`
- resizing the output image with `output width` and `output height`
- scaling the full source image by percentage
- optional safe mode to prevent overwriting an existing destination file
- output images with `300 DPI`

## Requirements

- PowerShell
- `System.Drawing`
- `C:\Users\vital\Repositories\Utilities\ColorWriter.psm1`

## Usage

The script defines a function named `CropImage`. A typical usage pattern is:

```powershell
. .\CropImage.ps1
CropImage -s <source> -d <destination> [options]
```

## Parameters

- `-s`, `-source`, `-src`: Source image path. Must exist.
- `-d`, `-dest`, `-dst`: Destination image path.
- `-x`: Crop start x-coordinate.
- `-y`: Crop start y-coordinate.
- `-w`, `-width`: Crop width.
- `-h`, `-height`: Crop height.
- `-ow`, `-outputwidth`: Output image width.
- `-oh`, `-outputheight`: Output image height.
- `-p`, `-percent`: Scale factor for the full source image.
- `-save`: Safe mode. If enabled, the script throws an error when the destination file already exists.

## Behavior

- If `width` and `height` are omitted, the full source image is used.
- If only one crop dimension is given, the other dimension is calculated from the source aspect ratio.
- If no output size is given, the cropped size is used as output size.
- If only one output dimension is given, the other output dimension is calculated from the crop aspect ratio.
- If `percent` is used, the full source image is scaled and written to the destination file.
- The output file is saved as JPEG.

## Examples

Crop a fixed region:

```powershell
. .\CropImage.ps1
CropImage -s "C:\Images\Input.jpg" -d "C:\Images\Output.jpg" -x 0 -y 0 -w 936 -h 2800
```

Crop and resize:

```powershell
. .\CropImage.ps1
CropImage -s "C:\Images\Input.jpg" -d "C:\Images\Output.jpg" -x 100 -y 200 -w 800 -h 1200 -ow 400 -oh 600
```

Scale the full image by percentage:

```powershell
. .\CropImage.ps1
CropImage -s "C:\Images\Input.jpg" -d "C:\Images\Output.jpg" -p 0.35
```

## Notes

- The script currently contains project-specific example calls at the bottom of `CropImage.ps1`.
- Because of those embedded calls, running the script directly may execute one of those examples instead of only exposing the function.
- If you want a cleaner reusable tool, those example calls should be moved into comments, a separate example file, or guarded behind a condition.
