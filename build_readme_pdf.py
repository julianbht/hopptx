"""Convert README.md to README.pdf."""

from markdown_pdf import MarkdownPdf, Section

pdf = MarkdownPdf(toc_level=2)
pdf.add_section(Section(open("README.md").read()))
pdf.save("README.pdf")
print("Created README.pdf")
