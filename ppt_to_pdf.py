"""Export the generated deck to PDF using the installed PowerPoint."""

import os
import sys

import comtypes.client

ROOT = os.path.dirname(os.path.abspath(__file__))
PPTX = os.path.join(ROOT, "Multi_Agent_Research_System_PPT_Suraj_Singh_Rawat.pptx")
PDF = os.path.join(ROOT, "Multi_Agent_Research_System_PPT_Suraj_Singh_Rawat.pdf")

FORMAT_PDF = 32  # ppSaveAsPDF


def main():
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    deck = powerpoint.Presentations.Open(PPTX, WithWindow=False)
    try:
        deck.SaveAs(PDF, FORMAT_PDF)
    finally:
        deck.Close()
        powerpoint.Quit()
    print("wrote", PDF)


if __name__ == "__main__":
    main()
