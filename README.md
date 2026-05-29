# Partnership Outreach Automation System

AI-powered email outreach automation built with 
Google Apps Script, Claude AI, and Gmail API.

## Overview

An end-to-end automation system that monitors 
funding calls and tender opportunities, matches 
them to partner profiles, generates personalised 
outreach emails via Claude AI, manages a human 
approval step, sends via Gmail, tracks replies, 
and handles follow-ups. All operating inside 
Google Sheets with no extra tools or 
subscriptions required.

## Demo

Watch the full system demo:
https://www.loom.com/share/772d0e2bd350433c821aaff5a48c0485

## Tech Stack

- Google Apps Script — automation engine
- Claude API (Anthropic) — AI email generation
- Gmail API — email sending and reply tracking
- Google Sheets — database and control panel

## Features

- AI-powered personalised email generation
- Human approval required before any send
- Automated follow-up sequence at 7 and 14 days
- Reply detection and tracking
- 30-day contact cooldown enforcement
- Duplicate detection across all opportunities
- Do Not Contact list enforcement
- Auto-closure after no response
- Full error logging with row references
- Zero additional paid subscriptions

## System Flow
[Incoming Tab] > [validateIncoming()]

[Calls for Proposals] > [Claude AI]
[Human Approval] > [Gmail Send]
[Reply Tracking] > [Follow-up Sequence]

## Sheet Structure

| Tab | Purpose |
|-----|---------|
| Partners | Partner profiles and contact history |
| Calls for Proposals | Main working tab |
| Incoming | Scraper buffer with validation |
| Do Not Contact | Blacklist enforcement |
| Errors | Automatic error logging |
| Settings | Configurable without code changes |

## Automation Triggers

| Function | Schedule |
|----------|----------|
| validateIncoming | Every 1 hour |
| procesNewOpportunities | Every 30 minutes |
| sendApprovedEmails | Every 15 minutes |
| checkForReplies | Daily 8am |
| processFollowUp1 | Daily 9am |
| processFollowUp2 | Daily 10am |

## File Structure

| File | Purpose |
|------|---------|
| Code.gs | Main processing logic |
| Claude.gs | Claude API integration |
| Gmail.gs | Gmail sending and reply tracking |
| Config.gs | Settings management |
| Utils.gs | Helper functions and error logging |
| Incoming.gs | Incoming validation and processing |

## How It Works

1. Scraper deposits new opportunities 
   into Incoming tab
2. validateIncoming() checks for duplicates 
   and moves valid rows to Calls for Proposals
3. Claude AI generates personalised email 
   for each opportunity-partner match
4. Team reviews draft and marks Yes to approve
5. System sends via Gmail and logs Thread ID
6. System checks for replies daily
7. Follow-ups generated at day 7 and day 14
8. Rows auto-closed after second follow-up 
   with no response

## Configuration

All settings managed from Settings tab:
- Claude API Key
- Sender Name and Email
- Follow-up delay days
- Cooldown period
- Email prompt template
