# HCI Merged Records Processor

A Node.js script for processing and merging HCI (Healthcare Information) records across multiple Excel worksheets.

## Description

This tool processes Excel workbooks containing healthcare contact records, matching and merging records between two sheets:
- "EE Portal Merged Record IDs"
- "ISSA Portal"

The script identifies matching records and consolidates them, updating the merged contact IDs in the appropriate columns.

## Prerequisites

- Node.js (v12.0.0 or higher)
- npm (Node Package Manager)

## Installation

1. Clone the repository: