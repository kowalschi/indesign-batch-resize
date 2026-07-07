
# Adobe InDesign Batch Resize – ExtendScript/JSX Automation
 
A collection of ExtendScript (JSX) automation tools for Adobe InDesign, developed and iterated over **6 years (2020–2026)**. These scripts automate repetitive resizing and document preparation tasks for advertising/print production, evolving from early single-purpose versions to multi-master, cross-platform (Win/Mac) solutions supporting multiple output formats.
 
## Highlights
 
- **2020** – Initial resize scripts (v01–v09b), establishing core automation logic
- **2023** – Introduced multi-master page support (v10a, v11_IQOS)
- **2024** – Cross-platform compatibility (Win/Mac), multi-variant outputs (v11–v12)
- **2025–2026** – Refined download/export automation, latest multi-master versions (v14–v18)
This repository documents the progression of my JavaScript/ExtendScript skills applied to real production workflows in graphic design automation.


# Adobe InDesign Batch Resize Script

I developed a script for Adobe InDesign that automates the resizing process. In advertising, it is often necessary to take a key visual and produce multiple variants of it — same content, different dimensions — for print or web use. This process is commonly referred to as **resizing** or **declination** of a key visual. The script can handle multiple key visuals at once and resize them to any required dimensions.

The script takes its input from an Excel file containing the new file names, target sizes, and additional information such as the output location of print materials or the destination format (web or print).

[Visit my site at resize dot magic dot re] 


### How it works

The method used by the script relies on a **master Adobe InDesign document** that contains multiple page variants — portrait, square, and landscape. The script reads the Excel file line by line and uses this master file to generate new InDesign files at the specified dimensions. It selects the appropriate page from the master based on the orientation of the target format: portrait pages are used as the basis for new portrait sizes, landscape pages for landscape outputs, and so on.

The determining factor for page selection is the **aspect ratio** (width divided by height). The master file contains a comprehensive set of ratio-based pages, designed to cover virtually all possible aspect ratios a new key visual might require.

### Label-based positioning rules

The script uses a **labeling system** to define precise positioning rules for each element on the page. For example:
- The **logo** must always be placed in the bottom-right corner at defined distances from the margins.
- The **background** must fully cover the surface of the key visual, including bleed for print.
- **Legal or warning text** must always appear at the bottom, with a height equal to 10% of the visual's total height.
- The **"New" badge** must be positioned in the top-left corner with defined spacing from the margins.
- A **transparent gradient overlay** must sit above the background at the bottom of the page and must include bleed.

The script applies all of these rules automatically, honoring each label's defined behavior during the resizing process.

### Impact

After the script completes, human intervention is drastically reduced. Manually resizing a key visual to a new format typically takes between 30 minutes and several hours. With the script, the same task requires approximately **5 minutes of human involvement** per format — and the time savings multiply significantly when hundreds of new sizes need to be produced.

### Who it helps

- Marketing teams
- Brand designers
- Print production staff
- Anyone who needs one campaign design in many poster sizes

### Why it matters

- Saves time by avoiding manual resizing for each format
- Ensures consistent visuals across all poster sizes
- Converts one master layout into many dimensioned versions automatically

### Simple way to explain it

- “This tool turns one main poster design into many different poster sizes automatically.”
- “It’s like making a batch of resized ads from the same base artwork.”
- “You build the key visual once, then the script makes all the other formats.”
