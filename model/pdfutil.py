"""HTML → PDF через установленный Chromium (headless), альбомная A4."""
import subprocess, os, tempfile
def html_to_pdf(html_path, pdf_path):
    css = "@page{size:A4 landscape;margin:8mm}"
    s = open(html_path).read().replace("<style>", f"<style>{css}", 1)
    tmp = html_path.replace(".html", "_p.html"); open(tmp, "w").write(s)
    subprocess.run(["/opt/pw-browsers/chromium", "--headless", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
                    f"--print-to-pdf={os.path.abspath(pdf_path)}", "file://" + tmp], check=True, capture_output=True, timeout=180)
