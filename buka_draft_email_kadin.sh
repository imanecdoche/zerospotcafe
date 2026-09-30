#!/usr/bin/env bash
PDF_PATH="/media/fatihfarhat/New Volume1/FATIH DATA/ZeroSpot Cafe/PROPOSAL_INVESTASI_ZEROSPOT_CAFE.pdf"
if [ ! -f "$PDF_PATH" ]; then
    PDF_PATH="/home/fatihfarhat/ZeroSpot_Backup/PROPOSAL_INVESTASI_ZEROSPOT_CAFE.pdf"
fi

thunderbird -compose "to='kadinbanten@gmail.com',subject='Penawaran Kemitraan & Proposal Investasi UMKM: ZeroSpot Cafe Cikedal (Kopi Lokal Banten)',attachment='$PDF_PATH'"
