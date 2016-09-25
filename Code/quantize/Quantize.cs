using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Collections;

public unsafe class Quantizer {

	private Hashtable _colorMap;	// lookup table for colors
	private Color[]	_colors;		// list of all colors in the palette

	/*---COMMENT---------------------------------------------------------------
	Construct the quantizer

	Date:	Name:	Description:
	5/?/03	MS		http://msdn.microsoft.com/library/en-us/dnaspp/html/colorquant.asp
	-------------------------------------------------------------------------*/
	public Quantizer(ArrayList palette) {
		_colorMap = new Hashtable();
		_colors = new Color[palette.Count];
		palette.CopyTo(_colors);
	}

	/*---COMMENT---------------------------------------------------------------
	Quantize an image and return the resulting output bitmap

	Date:	Name:	Description:
	5/?/03	MS		http://msdn.microsoft.com/library/en-us/dnaspp/html/colorquant.asp
	-------------------------------------------------------------------------*/
	public Bitmap Quantize(Image source) {
		int	height = source.Height;
		int width = source.Width;

		Rectangle bounds = new Rectangle(0, 0, width, height);
		Bitmap copy = new Bitmap(width, height, PixelFormat.Format32bppArgb);
		Bitmap output = new Bitmap(width, height, PixelFormat.Format8bppIndexed);

		// draw the source image onto the copy
		using(Graphics g = Graphics.FromImage(copy)) {
			g.PageUnit = GraphicsUnit.Pixel;
			g.DrawImageUnscaled(source, bounds);
		}
		BitmapData sourceData = null;		// pointer to bitmap data

		try {
			sourceData = copy.LockBits(bounds, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
			output.Palette = this.GetPalette(output.Palette);
			ConvertPalette(sourceData, output, width, height, bounds);
		} finally {
			copy.UnlockBits(sourceData);
		}
		return output;
	}

	/*---COMMENT---------------------------------------------------------------
	Execute a second pass through the bitmap

	Date:	Name:	Description:
	5/?/03	MS		http://msdn.microsoft.com/library/en-us/dnaspp/html/colorquant.asp
	-------------------------------------------------------------------------*/
	private void ConvertPalette(BitmapData sourceData, Bitmap output,
		int width, int height, Rectangle bounds) {

		BitmapData outputData = null;

		try {
			outputData = output.LockBits(bounds, ImageLockMode.WriteOnly, PixelFormat.Format8bppIndexed);

			// Define the source data pointers. The source row is a byte to
			// keep addition of the stride value easier(as this is in bytes)
			byte*	pSourceRow = (byte*)sourceData.Scan0.ToPointer();
			Int32*	pSourcePixel = (Int32*)pSourceRow;
			Int32*	pPreviousPixel = pSourcePixel;

			// destination data pointers
			byte*	pDestinationRow = (byte*)outputData.Scan0.ToPointer();
			byte*	pDestinationPixel = pDestinationRow;

			// convert first pixel to prime the loop
			byte	pixelValue = QuantizePixel((Color32*)pSourcePixel);

			// assign the value of the first pixel
			*pDestinationPixel = pixelValue;

			// loop through each pixel row
			for (int row = 0; row < height; row++,
				pSourceRow += sourceData.Stride, pDestinationRow += outputData.Stride) {

				// set pixels to first in row
				pSourcePixel = (Int32*)pSourceRow;
				pDestinationPixel = pDestinationRow;

				// loop through each pixel on this scan line
				for (int col = 0; col < width; col++, pSourcePixel++, pDestinationPixel++) {
					if (*pPreviousPixel != *pSourcePixel) {
						// only get new value if different from last pixel
						pixelValue = QuantizePixel((Color32*)pSourcePixel);
						pPreviousPixel = pSourcePixel;
					}
					*pDestinationPixel = pixelValue;
				}
			}
		} finally {
			output.UnlockBits(outputData);
		}
	}

	/*---COMMENT---------------------------------------------------------------
	Process an individual pixel

	Date:	Name:	Description:
	5/?/03	MS		http://msdn.microsoft.com/library/en-us/dnaspp/html/colorquant.asp
	-------------------------------------------------------------------------*/
	private byte QuantizePixel(Color32* pixel) {
		byte	colorIndex = 0;
		int		colorHash = pixel->ARGB;	

		if(_colorMap.ContainsKey(colorHash)) {
			// color found in lookup table
			colorIndex = (byte)_colorMap[colorHash];
		} else {
			// not found so loop through palette to find the nearest match.
			if(0 == pixel->Alpha) {
				// if alpha is 0 lookup transparent color value
				for(int index = 0; index < _colors.Length; index++) {
					if(0 == _colors[index].A) { colorIndex = (byte)index; break; }
				}
			} else {
				// not transparent
				int	leastDistance = int.MaxValue;
				int red = pixel->Red;
				int green = pixel->Green;
				int blue = pixel->Blue;

				// loop through palette to find closest match
				for(int index = 0; index < _colors.Length; index++) {
					Color paletteColor = _colors[index];
					
					int	redDistance = paletteColor.R - red;
					int	greenDistance = paletteColor.G - green;
					int	blueDistance = paletteColor.B - blue;

					int	distance = (redDistance * redDistance) + 
						(greenDistance * greenDistance) + 
						(blueDistance * blueDistance);

					if(distance < leastDistance) {
						colorIndex = (byte)index;
						leastDistance = distance;

						// exact match found so exit loop
						if(0 == distance) { break; }
					}
				}
			}
			// Now I have the color, pop it into the hashtable for next time
			_colorMap.Add(colorHash, colorIndex);
		}

		return colorIndex;
	}


	/*---COMMENT---------------------------------------------------------------
	transfer palette from passed array list, color array

	Date:	Name:	Description:
	5/?/03	MS		http://msdn.microsoft.com/library/en-us/dnaspp/html/colorquant.asp
	-------------------------------------------------------------------------*/
	private ColorPalette GetPalette(ColorPalette palette) {
		for(int index = 0; index < _colors.Length; index++) {
			palette.Entries[index] = _colors[index];
		}
		// make first color (background) transparent
		palette.Entries[0] = Color.FromArgb(0, _colors[0]);
		return palette;
	}

	/*---COMMENT---------------------------------------------------------------
	Struct that defines a 32 bpp colour
	This struct is used to read data from a 32 bits per pixel image
	in memory, and is ordered in this manner as this is the way that
	the data is layed out in memory

	Date:	Name:	Description:
	5/?/03	MS		http://msdn.microsoft.com/library/en-us/dnaspp/html/colorquant.asp
	-------------------------------------------------------------------------*/
	[StructLayout(LayoutKind.Explicit)]
	public struct Color32 {
		[FieldOffset(0)] public byte Blue;
		[FieldOffset(1)] public byte Green;
		[FieldOffset(2)] public byte Red;
		[FieldOffset(3)] public byte Alpha;
		[FieldOffset(0)] public int ARGB;

		public Color Color {
			get	{ return Color.FromArgb(Alpha, Red, Green, Blue); }
		}
	}
}