"""Builds conferences_2027.xlsx, CSVs and data.json from the CONFS list below.
Edit the list, then run: python3 build.py"""
import csv, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

CHECKED = "2026-11-17"
COLS = ["Group", "Acronym", "Full name", "Edition", "Publisher / series", "Indexing",
        "Location", "Region", "Conference dates", "Submission deadline", "Deadline status",
        "Notification", "Camera-ready", "Mode", "Topic area", "Page limit",
        "Reg. fee (approx.)", "Fee: student", "Fee: non-IEEE / regular", "Fee: online/hybrid", "Website", "Checked on", "Notes",
        "Verification", "Track record", "Acceptance signals", "Fit score"]

def c(group, acr, name, ed, pub, idx, loc, reg, dates, dl, dls, notif, cr, mode, topic, pages, fee, url, notes,
      fee_student="Not listed", fee_regular="Not listed", fee_online="Not listed",
      verification="Listing site", track="Not checked", accept="Not published"):
    return dict(zip(COLS, [group, acr, name, ed, pub, idx, loc, reg, dates, dl, dls, notif, cr,
                           mode, topic, pages, fee, fee_student, fee_regular, fee_online, url, CHECKED, notes,
                           verification, track, accept, 0]))

CONFS = [
    # ---------- IEEE ----------
    c("IEEE", "ICETICS 2027", "Intl. Conf. on Emerging Trends in Information, Communication & Systems", "2nd",
      "IEEE (record #74017; tech. sponsor: IEEE MP Section)", "IEEE Xplore (expected)", "Bhopal, MP", "Central",
      "23–24 Jul 2027", "2027-01-10", "Confirmed", "From 15 Mar 2027", "30 Apr 2027", "Hybrid",
      "AI, data science, intelligent systems track", "6 pages, IEEE 2-col (extra pages ₹500 each)", "₹5,310–7,080 incl. GST by category/mode",
      "https://bgibhopal.com/icetics-2027/", "Host: Bansal Institute of Science & Technology. Best fit in IEEE list. IEEE record #74017 and hybrid mode shown on official site (16 Nov 2026); the dates sub-page could not be re-opened this run, so the deadline is carried over from the earlier check.",
      fee_student="₹6,490 online / ₹5,310 offline non-IEEE (IEEE: ₹5,900 / ₹4,720), incl. 18% GST",
      fee_regular="₹7,080 online / ₹6,490 offline non-IEEE academician (IEEE: ₹6,490 / ₹5,310), incl. 18% GST",
      fee_online="Online rates are listed and are higher than offline as shown on the site: ₹6,490 student, ₹7,080 academician (non-IEEE), incl. 18% GST"),
    c("IEEE", "ICCCA 2027", "IEEE Intl. Conf. on Computing, Communication and Automation", "8th",
      "IEEE (record #72457; IEEE UP Section & IAS sponsored)", "IEEE Xplore (expected)", "Greater Noida, UP", "North",
      "10–12 Apr 2027", "2027-02-15", "Confirmed", "10 Mar 2027", "15 Mar 2027", "Hybrid",
      "Computing, AI, communication", "6 pages, IEEE 2-col (+2 extra pages at ₹1000/page)",
      "₹10,620 non-IEEE / ₹9,500 IEEE member (early bird to 15 Mar 2027, incl. GST); student author ₹8,260 / ₹7,080",
      "https://www.iccca.co.in/", "Galgotias University. Only 20% of presentations online, so an online slot is not guaranteed.",
      fee_student="₹8,260 non-IEEE early bird to 15 Mar 2027, ₹8,850 after (IEEE: ₹7,080 / ₹8,260), incl. 18% GST",
      fee_regular="₹10,620 non-IEEE early bird, ₹11,800 after (IEEE: ₹9,500 / ₹10,620), incl. 18% GST",
      fee_online="Not listed (no separate online rate)"),
    c("IEEE", "ICCMC 2027", "Intl. Conf. on Computing Methodologies and Communication", "10th",
      "IEEE (past editions in Xplore)", "IEEE Xplore (2027 unconfirmed)", "Erode, TN", "South",
      "21–23 Jun 2027", "2027-03-29", "Confirmed", "26 Apr 2027", "Not listed (reg. by 24 May 2027)", "Unconfirmed",
      "Computing, AI/ML, communication", "Not listed", "Not listed (site: \"to be updated\")",
      "https://icocmc.com/2027/", "Surya Engineering College. Site still shows no 2027 IEEE record number; confirm before submitting."),
    c("IEEE", "AIC 2027", "IEEE World Conf. on Applied Intelligence and Computing", "6th",
      "IEEE (proceedings \"will be submitted for inclusion to IEEE Xplore\"; no record no. shown)", "IEEE Xplore (expected)", "SRIT, Jabalpur, MP", "Central",
      "26–27 Jun 2027", "2027-04-17", "Confirmed", "27 May 2027", "Not listed (reg. by 11 Jun 2027)", "Unconfirmed",
      "Applied AI, computing", "6 pages IEEE format before extra-page charges (₹2,000/page)", "₹12,000 + GST (₹10,000 SCRS/IEEE members)",
      "https://www.scrs.in/conference/aic2027", "SCRS index and conference page now agree on 26–27 Jun 2027 (earlier date conflict resolved).",
      fee_student="No student rate listed",
      fee_regular="₹12,000 + GST (₹10,000 + GST for SCRS/IEEE members)",
      fee_online="Same fee; it covers online or offline presentation (offline on Day 1 only)"),
    # ---------- Springer ----------
    c("Springer", "PEIS 2027", "Intl. Conf. on Power Engineering and Intelligent Systems", "5th",
      "Springer LNEE (\"accepted papers will be submitted to\")", "Scopus", "NIT Uttarakhand", "North",
      "13–14 Mar 2027", "2027-01-02", "Confirmed", "11 Feb 2027", "Not listed (reg. by 26 Feb 2027)", "Hybrid",
      "Intelligent systems, power + AI", "Not listed", "Not listed",
      "https://scrs.in/conference/peis2027", "SCRS-organized. Leans power engineering; AI papers fit the intelligent-systems track.",
      fee_student="Not listed (fee page is blank)",
      fee_regular="Not listed (fee page is blank)",
      fee_online="Not listed (fee page is blank)"),
    c("Springer", "BIDA 2027", "Intl. Conf. on Business Intelligence & Data Analytics", "4th",
      "Springer \"Studies in Autonomic, Data-driven and Industrial Computing\" (approval pending)", "Scopus (if approved)", "RVIM, Bengaluru, KA", "South",
      "10–11 Apr 2027", "2027-01-30", "Confirmed", "11 Mar 2027", "Not listed (reg. by 26 Mar 2027)", "Hybrid",
      "Data analytics, business intelligence, applied ML", "Not listed", "₹9,000 + 18% GST (₹8,000 SCRS/host members)",
      "https://scrs.in/conference/bida2027", "NEW 16 Nov 2026. SCRS-organized; management-school host, so it suits applied analytics/ML papers more than core DL/CV.",
      fee_student="No student rate listed",
      fee_regular="₹9,000 + 18% GST (₹8,000 + GST for SCRS/host members)",
      fee_online="Same fee; it covers online or offline presentation (offline on Day 1 only)"),
    c("Springer", "CVR 2027", "Intl. Conf. on Computer Vision and Robotics", "7th",
      "Springer LNNS (approval pending)", "Scopus", "NIT Goa", "West",
      "24–25 Apr 2027", "2027-02-13", "Confirmed", "25 Mar 2027", "Not listed (reg. by 9 Apr 2027)", "Hybrid",
      "Computer vision, robotics, DL", "Not listed", "₹9,000 + 18% GST (₹8,000 SCRS/NIT Goa members)",
      "https://scrs.in/conference/cvr2027", "SCRS-organized. Best CV fit. Offline sessions on Day 1 only; fee covers online or offline presentation.",
      fee_student="No student rate listed",
      fee_regular="₹9,000 + 18% GST (₹8,000 + GST for SCRS/host members)",
      fee_online="Same fee; it covers online or offline presentation (offline on Day 1 only)"),
    c("Springer", "CIMA 2027", "Congress on Intelligent Machines and Algorithms", "3rd",
      "Springer LNNS (approval pending)", "Scopus", "NIT Puducherry", "South",
      "22–23 May 2027", "2027-03-13", "Confirmed", "Not listed", "Not listed", "Hybrid",
      "Computer vision, AI in healthcare, robotics, IoT, intelligent systems", "12 pages Springer format before extra-page charges (₹500/page)",
      "₹9,000 + 18% GST (₹8,000 SCRS/NIT Puducherry members)",
      "https://scrs.in/conference/cima2027", "Replaces the earlier \"CVR 2027 (Puducherry)\" row: the SCRS index entry at NIT Puducherry with a 13 Mar deadline is CIMA, not a second CVR.",
      fee_student="No student rate listed",
      fee_regular="₹9,000 + 18% GST (₹8,000 + GST for SCRS/host members)",
      fee_online="Same fee; it covers online or offline presentation (offline on Day 1 only)"),
    c("Springer", "ICIVC 2027", "Intl. Conf. on Intelligent Vision and Computing", "7th",
      "Springer LNNS (approval pending)", "Scopus", "ICFAI University, Dehradun, UK", "North",
      "11–12 Jun 2027", "2027-04-02", "Confirmed", "12 May 2027", "Not listed (reg. by 27 May 2027)", "Hybrid",
      "Computer vision, intelligent computing, DL", "Not listed", "₹9,000 + 18% GST (₹8,000 SCRS/ICFAI members)",
      "https://scrs.in/conference/icivc2027", "NEW 16 Nov 2026. SCRS-organized; strong CV fit. Offline sessions on Day 1 only.",
      fee_student="No student rate listed",
      fee_regular="₹9,000 + 18% GST (₹8,000 + GST for SCRS/host members)",
      fee_online="Same fee; it covers online or offline presentation (offline on Day 1 only)"),
    c("Springer", "ICDWCI 2027", "Intl. Conf. on Distributed Wireless Communication and IoT", "3rd",
      "Springer (planned, IoT & Wireless Comm. Networks series)", "Not confirmed", "Cauvery Inst. of Technology, Mandya, KA", "South",
      "28–29 May 2027", "2027-04-15", "Confirmed", "30 Apr 2027", "Not listed (reg. by 15 May 2027)", "Unconfirmed",
      "Wireless, IoT; AI/ML/data mining listed as allied areas", "8 pages before extra-page charges (₹500/page)", "₹4,000 academician", "https://icdwci.co.in/",
      "Full name corrected from the official site (wireless/IoT focus, weaker AI fit). Springer publication only planned so far.",
      fee_student="Not listed",
      fee_regular="₹4,000 academician, ₹5,000 industry (GST not stated)",
      fee_online="Not listed"),
    c("Springer", "ICDSA 2027", "Intl. Conf. on Data Science and Applications", "8th",
      "Springer LNNS (approval pending)", "Scopus", "VIT Mauritius (+ online)", "Outside India",
      "7–8 Aug 2027", "2027-05-29", "Confirmed", "8 Jul 2027", "Not listed (reg. by 23 Jul 2027)", "Hybrid",
      "Data science, ML, DL", "Not listed", "₹9,000 + 18% GST (₹8,000 SCRS/host members)",
      "https://scrs.in/conference/icdsa2027", "Dates moved on the official site: conference 7–8 Aug 2027 (was 11–12 Jul), deadline 29 May 2027 (was 2 May). Host is in Mauritius, but online presentation lets you attend from India.",
      fee_student="No student rate listed",
      fee_regular="₹9,000 + 18% GST (₹8,000 + GST for SCRS/host members)",
      fee_online="Same fee; it covers online or offline presentation (offline on Day 1 only)"),
    # ---------- Publisher to confirm ----------
    c("Publisher TBC", "ICAN 2027", "Intl. Conf. on AI-driven Computing, Analytics and Networks", "7th",
      "Not announced (\"Scopus-indexed conference proceedings\"; publication partner \"to be updated soon\")", "Scopus (per site)", "Chitkara Univ., Chandigarh/Rajpura, PB", "North",
      "16–17 Jul 2027", "2027-03-15", "Confirmed", "15 May 2027", "Not listed (reg. by 11 Jun 2027)", "In-person for India-affiliated authors (virtual only for authors outside India)",
      "AI, analytics, networks", "4,000–6,000 words, format-free initial submission", "₹9,000 student / ₹12,000 faculty (non-IEEE)",
      "https://www.chitkara.edu.in/cse-can/", "Strong topic fit and dates now confirmed on the official site, but Indian authors must present physically. Submissions open 10 Jan 2027.",
      fee_student="₹9,000 non-IEEE, ₹8,000 IEEE (GST not stated)",
      fee_regular="₹12,000 non-IEEE faculty/researcher, ₹10,000 IEEE (GST not stated)",
      fee_online="No online option for India-affiliated authors (virtual only for authors outside India)"),
    c("Publisher TBC", "ICETCI 2027", "Intl. Conf. on Emerging Technologies and Computing Innovations", "3rd",
      "Springer proposal submitted (2025 and 2026 proceedings published with Springer)", "Not confirmed", "Mumbai, MH", "West",
      "20–21 Feb 2027", "2027-01-10", "Confirmed", "20 Jan 2027", "Not listed", "Hybrid",
      "Computing, emerging tech", "Not listed", "₹11,000 student/scholar; ₹12,000 faculty; ₹13,000 industry",
      "https://www.icetci.in/", "Organized by MG Aricent Educational Foundation (Nashik). Site states publication and indexing are not guaranteed; check publisher carefully.",
      fee_student="₹11,000 student/research scholar (GST not stated)",
      fee_regular="₹12,000 faculty, ₹13,000 industry (GST not stated)",
      fee_online="Not listed (no separate online rate)"),
    # ---------- Watchlist (outside window) ----------
    c("Watchlist", "NEXAI 2027", "Next-Generation AI & Agentic Intelligence", "—",
      "Springer LNNS", "Scopus", "NorthCap Univ., Gurugram, HR", "North",
      "4–5 Mar 2027", "2026-11-22", "Outside window", "1 Jan 2027 (rolling)", "25 Jan 2027", "Hybrid",
      "AI, agentic AI, GenAI/LLMs, NLP", "", "", "https://nexai2027.ncuindia.edu/",
      "Deadline now published: 22 Nov 2026, before the window. Watch for an extension into January.",
      fee_student="₹7,500 early bird to 25 Jan 2027, ₹8,500 regular (GST not stated)",
      fee_regular="₹9,000 early bird, ₹11,000 regular, academician/faculty (GST not stated)",
      fee_online="Not listed (no separate online rate)"),
    c("Watchlist", "ICCCT 2027", "Intl. Conf. on Communication & Computational Technologies", "9th",
      "Springer LNNS (approval pending)", "Scopus", "NFSU Goa Campus", "West",
      "30–31 Jan 2027", "2026-11-21", "Outside window", "", "", "Hybrid",
      "Communication, computational technologies", "", "", "https://scrs.in/conference/iccct2027",
      "NEW 16 Nov 2026 (watch only). Deadline from the SCRS conference index; before window.",
      fee_student="No student rate listed",
      fee_regular="₹9,000 + 18% GST (₹8,000 + GST for SCRS/host members)",
      fee_online="Same fee; it covers online or offline presentation (offline on Day 1 only)"),
    c("Watchlist", "ICRCICN 2027", "IEEE Intl. Conf. on Research in Computational Intelligence and Comm. Networks", "8th",
      "IEEE (record #73691; IEEE Kolkata Section tech. co-sponsor)", "IEEE Xplore", "Kalyani, WB", "East",
      "27–28 Feb 2027", "2026-11-30", "Outside window", "1 Feb 2027", "14 Feb 2027", "Hybrid",
      "Computational intelligence, networks", "", "", "https://icrcicn.in/", "Deadline before window, no extension announced as of 16 Nov 2026; extensions common.",
      fee_student="₹8,500 non-IEEE early bird, ₹10,000 late (IEEE: ₹8,000 / ₹9,500); GST not stated",
      fee_regular="₹10,000 non-IEEE academic early bird, ₹11,000 late (IEEE: ₹9,000 / ₹10,000); GST not stated",
      fee_online="Not listed (no separate online rate)"),
    c("Watchlist", "SANKALP 2027", "SANKALP 2027 (ABV-IIITM Gwalior IEEE conference)", "—",
      "IEEE (per search listing)", "Not verified", "ABV-IIITM, Gwalior, MP", "Central",
      "9–11 Apr 2027", "2026-11-30", "Outside window", "", "15 Feb 2027", "Unconfirmed",
      "Computing, AI", "", "", "https://sankalp.iiitm.ac.in/",
      "NEW 16 Nov 2026 (watch only). Official site would not load (expired TLS certificate); dates come from a search-result snippet of that site and are unverified."),
    c("Watchlist", "INDIACom 2027", "Intl. Conf. on Computing for Sustainable Global Development", "21st INDIACom / 14th intl.",
      "IEEE (Conference ID 74162; Delhi Section tech. approved)", "IEEE Xplore, Scopus", "New Delhi", "North",
      "7–9 Apr 2027", "2026-12-07", "Outside window", "8 Feb 2027", "22 Feb 2027", "Hybrid-friendly",
      "Computing, AI", "", "", "https://bvicam.ac.in/indiacom/", "Deadline before window, no extension announced as of 16 Nov 2026; watch for one.",
      fee_student="₹8,000 general early bird to 22 Feb 2027, ₹9,000 after (society members: ₹6,400 / ₹7,200); GST not stated",
      fee_regular="₹10,000 general teacher/research scholar early bird, ₹11,000 after (members: ₹8,000 / ₹8,800); GST not stated",
      fee_online="Not listed (no separate online rate)"),
    c("Watchlist", "IEM-ICDC 2027", "Intl. Conf. on Computational Intelligence, Data Science and Cloud Computing", "5th",
      "Springer format (series not stated)", "Not verified", "IEM, Kolkata, WB", "East",
      "17–19 Mar 2027", "2026-12-10", "Outside window", "", "", "Unconfirmed",
      "Computational intelligence, data science", "Minimum 10 pages, Springer format", "", "https://www.iemicdc.org/",
      "NEW 16 Nov 2026 (watch only). Source: EasyChair CFP page, not the conference site."),
    c("Watchlist", "CONECCT 2027", "IEEE Intl. Conf. on Electronics, Computing and Communication Technologies", "13th",
      "IEEE (Bangalore Section)", "IEEE Xplore", "Bengaluru, KA", "South",
      "29–31 Jul 2027", "2026-12-15", "Outside window", "15 Feb 2027", "", "Unconfirmed",
      "18 tracks in 7 focus areas; AI/ML embedded in tracks, not a headline track", "", "",
      "https://ieee-conecct.org/", "CORRECTED 16 Nov 2026: official site gives 15 Dec 2026 as the paper deadline; 15 Feb 2027 is the acceptance notification, not the deadline. Moved from IEEE to Watchlist.",
      fee_student="Not listed (registration page returns 404)",
      fee_regular="Not listed (registration page returns 404)",
      fee_online="Not listed (registration page returns 404)"),
    c("Watchlist", "IIPS-CoMET 2027", "IEEE Intl. Conf. on Modern AI, Entrepreneurship & Technology", "—",
      "IEEE (record #73825)", "IEEE Xplore (expected)", "IIPS, DAVV, Indore, MP", "Central",
      "19–20 Mar 2027", "2026-12-21", "Outside window", "", "", "Hybrid",
      "Human-centric AI, AI in business", "", "", "https://iips.edu.in/comet27/",
      "NEW 16 Nov 2026 (watch only). Deadline 10 days before the window; an extension would bring it in.",
      fee_student="Not listed (registration \"coming soon\")",
      fee_regular="Not listed (registration \"coming soon\")",
      fee_online="Not listed (registration \"coming soon\")"),
    c("Publisher TBC", "ISC 2027", "Intl. Conf. on Intelligent Systems and Computing (SCRS)", "—",
      "Not stated on conference page", "Not confirmed", "SVVV, Indore, MP (SCRS index says \"Australia\")", "Central",
      "11–12 Sep 2027", "2027-07-03", "Confirmed", "12 Aug 2027", "Not listed (reg. by 27 Aug 2027)", "Unconfirmed",
      "Intelligent systems, AI/ML", "Not listed", "Not listed", "https://scrs.in/conference/isc2027",
      "Rescheduled from 15–16 May to 11–12 Sep 2027; deadline 3 Jul 2027. Venue conflict between the conference page (SVVV) and the SCRS index (Australia)."),
    # ---------- Added 17 Nov 2026: Jun–Aug 2027 deadlines ----------
    c("Springer", "CIS 2027", "Congress on Intelligent Systems", "8th",
      "Springer LNNS (approval pending; earlier editions in LNNS)", "Scopus", "NIT Warangal, TS", "South",
      "5–6 Sep 2027", "2027-06-27", "Confirmed", "6 Aug 2027", "Not listed", "Hybrid",
      "Intelligent systems, ML, DL, CV, NLP", "Not listed", "Not listed",
      "https://scrs.in/conference/cis2027", "SCRS-organized. Page also shows a conflicting 6 Sep 2027 'submission last date'; 27 Jun is the full-paper deadline. Fee likely the standard SCRS ₹9,000 + GST but not yet confirmed on this page.",
      "No student rate listed", "Not listed", "Not listed"),
    c("Publisher TBC", "ICST 2027", "Intl. Conf. on Intelligent Computing and Sustainable Technologies", "—",
      "Not stated (SCRS-listed; SCRS events usually go to Springer LNNS)", "Not confirmed", "IIT Patna, BR", "East",
      "21–22 Aug 2027", "2027-06-12", "Confirmed", "Not listed", "Not listed", "Unconfirmed",
      "Intelligent computing, AI", "Not listed", "Not listed",
      "https://scrs.in/conferences", "Only the SCRS index lists it so far; no conference page yet. Not the IEEE software-testing ICST (Spain)."),
    c("Publisher TBC", "ICSISCET 2027", "Intl. Conf. on Sustainable and Innovative Solutions for Current Challenges in Engg. & Technology", "9th",
      "Springer LNNS (approval pending)", "Scopus (if approved)", "Not listed (SCRS)", "Not listed",
      "30–31 Oct 2027", "2027-08-21", "Confirmed", "30 Sep 2027", "Not listed (reg. by 15 Oct 2027)", "Unconfirmed",
      "Engineering + AI applications (broad)", "Not listed", "Not listed",
      "https://scrs.in/conference/icsiscet2027", "Broad engineering scope; AI papers fit but it is not a CS-focused venue."),
    c("Springer", "ICRTC 2027", "Intl. Conf. on Recent Trends in Computing", "—",
      "Springer (series not stated)", "Not confirmed", "SRM IST, Ghaziabad, UP", "North",
      "2–3 Jul 2027", "Not published", "Not published", "Not listed", "Not listed", "Unconfirmed",
      "Computing, AI", "Not listed", "Not listed",
      "https://icrtc.in/", "Deadline not out yet; 2026 cycle suggests Mar–Apr 2027."),
    # ---------- Added 17 Nov 2026: recurring venues whose 2027 CFP is not out yet ----------
    c("Premier", "CVIP 2027", "Intl. Conf. on Computer Vision and Image Processing", "12th",
      "Springer LNCS/CCIS (2026 edition)", "Scopus", "Rotates (2026: NIT Calicut)", "Not announced",
      "Dec 2027 (expected)", "Expected Apr & Jul 2027", "Expected (2026 pattern)", "", "", "Unconfirmed",
      "Computer vision, image processing", "", "", "https://cvip2026.nitc.ac.in/",
      "Best-known India CV venue in Springer. 2026 had Round 1 15 Apr and Round 2 31 Jul (both extended). Watch for the 2027 site."),
    c("Premier", "INDICON 2027", "IEEE India Council Intl. Conference", "24th",
      "IEEE (India Council flagship)", "IEEE Xplore", "Rotates (2026: Chennai)", "Not announced",
      "Dec 2027 (expected)", "Expected Aug 2027", "Expected (2026 pattern)", "", "", "Unconfirmed",
      "All EE/CS tracks incl. AI/ML", "", "", "https://ieeeindicon.org/",
      "2026 deadline was 31 Aug. More selective than the other IEEE venues here."),
    c("Watchlist", "ICAIC 2027", "Intl. Conf. on Artificial Intelligence and Computing", "—",
      "Springer LNNS (2026 edition)", "Scopus", "Manipal University Jaipur, RJ", "North",
      "Oct 2027 (expected)", "Expected Aug 2027", "Expected (2026 pattern)", "", "", "Unconfirmed",
      "AI, computing", "", "", "https://icaic.in/",
      "2026: deadline 15 Aug, conference 22–24 Oct. Watch for the 2027 call."),
    c("Watchlist", "ICACCP 2027", "Intl. Conf. on Advanced Computational and Communication Paradigms", "—",
      "Springer (series not stated)", "Not confirmed", "SMIT, Sikkim", "East",
      "8–10 Apr 2027", "2026-12-20", "Outside window", "", "", "Unconfirmed",
      "Computing, communication, AI", "", "", "https://icaccpa.in/",
      "Deadline 20 Dec 2026, before the window; an extension would bring it in."),
    # ---------- Added 17 Nov 2026: user-requested checks ----------
    c("IEEE", "ICISS 2027", "Intl. Conf. on Intelligent Sustainable Systems", "9th",
      "IEEE (Conference ID #72087)", "IEEE Xplore (expected)", "SCAD College of Engg. & Tech., Tirunelveli, TN", "South",
      "Not listed (past editions: Mar/Apr)", "2027-01-09", "Confirmed", "Not listed", "Not listed", "Unconfirmed",
      "Intelligent systems, AI/ML, IoT", "Not listed", "Not listed",
      "https://www.icoiss.com/", "Found while checking ICAISS. IEEE ID shown on site; hybrid option not mentioned."),
    c("Publisher TBC", "ICAISS 2027", "Intl. Conf. on Augmented Intelligence and Sustainable Systems", "5th",
      "Not stated for 2027 (earlier editions in IEEE Xplore)", "Not confirmed", "CARE College of Engg., Tiruchirappalli, TN", "South",
      "10–12 Jun 2027", "2027-03-11", "Confirmed", "15 Apr 2027", "Not listed (reg. by 12 May 2027)", "Unconfirmed",
      "AI, sustainable systems", "Not listed", "Not listed",
      "https://icaiss.com/2027/", "User-requested. Site does not yet state IEEE sponsorship or hybrid mode for 2027."),
    c("IEEE", "CICN 2027", "IEEE Intl. Conf. on Computational Intelligence and Communication Networks", "19th",
      "IEEE CPS", "IEEE Xplore", "Oriental University, Indore, MP", "Central",
      "23–25 Apr 2027", "Not published", "Not published", "Not listed", "Not listed", "Hybrid",
      "Computational intelligence, networks", "Not listed", "Not listed",
      "https://cicn.in/", "Organized by MIR Labs with IEEE MP Section. Deadline not yet on site; 2026 cycle suggests Jan–Feb."),
    c("Watchlist", "COEP CON 2027", "IEEE COEP Tech Intl. Conference", "—",
      "IEEE", "IEEE Xplore", "COEP Tech University, Pune, MH", "West",
      "4–6 Jun 2027", "2026-12-15", "Outside window", "Mar 2027", "", "Unconfirmed",
      "Engineering + computing tracks", "", "", "https://ieeecoepcon.coeptech.ac.in/",
      "User-requested. Deadline 15 Dec 2026, before the window; watch for an extension."),
    c("Watchlist", "IEEE-GCON 2027", "IEEE Guwahati Section Conference", "—",
      "IEEE", "IEEE Xplore", "Sohra, Meghalaya", "North-East",
      "16–18 Jun 2027", "2026-12-15", "Outside window", "", "", "Unconfirmed",
      "EE/CS tracks incl. AI", "", "", "https://gcon2027.in/", "Deadline 15 Dec 2026, before the window."),
    c("IEEE", "ICARC 2027", "Intl. Conf. on Advanced Research in Computing", "7th",
      "IEEE (tech. co-sponsor: IEEE Sri Lanka Section)", "IEEE Xplore", "Sabaragamuwa Univ., Sri Lanka", "Outside India (hybrid)",
      "16–17 Feb 2027", "2027-01-06", "Confirmed", "21 Jan 2027", "21 Jan 2027", "Hybrid",
      "Computing broadly, incl. AI/ML", "Not listed", "Not listed",
      "https://www.icarc.lk/", "Outside India but hybrid. Short gap between notification and camera-ready."),
    c("Watchlist", "IntelliSys 2027", "Intelligent Systems Conference (SAI)", "13th",
      "Springer LNNS", "Scopus", "Amsterdam (virtual option)", "Outside India (hybrid)",
      "2–3 Sep 2027", "2026-12-01", "Outside window", "1 Jan 2027", "15 Jan 2027", "Hybrid",
      "AI, intelligent systems", "", "", "https://saiconference.com/IntelliSys",
      "Round 2 deadline 1 Dec 2026, before the window. Virtual presentation allowed (~$600). Later rounds sometimes open."),
    c("Watchlist", "Computing Conference 2027", "Computing Conference (SAI)", "15th",
      "Springer LNNS", "Scopus", "London (virtual option in past editions)", "Outside India (hybrid)",
      "8–9 Jul 2027", "2026-12-01", "Outside window", "15 Jan 2027", "31 Jan 2027", "Unconfirmed",
      "Computing, AI", "", "", "https://saiconference.com/Computing",
      "Round 2 deadline 1 Dec 2026; watch for a Round 3. Virtual option not yet confirmed for 2027."),
    # ---------- Premier India venues (more selective; listed for reference) ----------
    c("Premier", "CVMI 2027", "IEEE Intl. Conf. on Computer Vision and Machine Intelligence", "6th",
      "IEEE", "IEEE Xplore", "IIITDM Kancheepuram, Chennai, TN", "South",
      "17–19 Jun 2027", "2027-02-25", "Confirmed (Round 2)", "15 Apr 2027", "1 Jun 2027", "In-person only",
      "Computer vision, machine intelligence", "Not listed", "Not listed",
      "https://cvmi2027.iiitdm.ac.in/", "Round 1 closed 15 Nov 2026. In-person presentation required for Xplore inclusion. Fees tentative.",
      "₹11,800 early / ₹12,980 late (non-IEEE student author, tentative, incl. GST)", "Not listed (tentative table on site)", "No online option"),
    c("Premier", "PReMI 2027", "Intl. Conf. on Pattern Recognition and Machine Intelligence", "12th",
      "Springer LNCS (biennial; 2025 in LNCS)", "Scopus, DBLP", "Rotates (2025: IIT Delhi)", "Not announced",
      "Dec 2027 (expected)", "Expected Jul 2027", "Expected (2025 pattern)", "", "", "Unconfirmed",
      "Pattern recognition, ML, CV", "", "", "https://link.springer.com/conference/premi",
      "No 2027 call yet. Odd-year series; 2025 was 11–14 Dec at IIT Delhi."),
    c("Premier", "ICCCNT 2027", "IEEE Intl. Conf. on Computing, Communication and Networking Technologies", "18th",
      "IEEE", "IEEE Xplore", "Rotates (usually an IIT/NIT)", "Not announced",
      "Jun–Jul 2027 (expected)", "Expected Mar 2027", "Expected (past pattern)", "", "", "Unconfirmed",
      "Computing, AI/ML, networking", "", "", "https://icccnt.com/", "No 2027 call yet. Large IEEE venue, hybrid in recent years."),
]

# Verification: "Official site" (details read from the organizer's page), "Listing site" (aggregator/index only),
# "Estimated" (no 2027 call yet; dates inferred from the last cycle).
# Track record: where past editions actually appeared. Acceptance signals: whatever the series publishes.
QUALITY = {
 "ICETICS 2027": ("Official site", "1st edition (2026) is in IEEE Xplore (conf. #69505).", "Not published"),
 "ICCCA 2027":   ("Official site", "In IEEE Xplore every year since 2015.", "2025: 2,278 submissions, 430 accepted (~19%), per organizer press release."),
 "ICCMC 2027":   ("Official site", "Editions 2017–2025 in IEEE Xplore; 2027 record number not yet shown.", "Large volumes (1,000+ papers/yr in recent editions); not selective."),
 "AIC 2027":     ("Official site", "SCRS series; earlier AIC editions in IEEE Xplore.", "Not published"),
 "ICISS 2027":   ("Official site", "2025 (7th) proceedings published by IEEE (3 vols).", "Large 3-volume proceedings; not selective."),
 "CICN 2027":    ("Official site", "In IEEE Xplore (CPS) since 2009.", "Not published"),
 "PEIS 2027":    ("Official site", "2026 edition promised Springer LNEE; volume not yet seen.", "Not published"),
 "BIDA 2027":    ("Official site", "SCRS series; prior BIDA editions in Springer.", "Not published"),
 "CVR 2027":     ("Official site", "LNNS proceedings volumes exist for 2021–2025.", "Not published"),
 "CIMA 2027":    ("Official site", "SCRS series; 2025 edition in Springer LNNS.", "Not published"),
 "ICIVC 2027":   ("Official site", "SCRS series; prior ICIVC editions in Springer LNNS.", "Not published"),
 "ICDWCI 2027":  ("Official site", "ICDWCI 2024 book published by Springer (IoT & Wireless Comm. Networks series); 2026 volume not yet seen.", "Not published"),
 "ICDSA 2027":   ("Official site", "2025 edition: 7 LNNS volumes; 2024 went to SN Computer Science special issue instead.", "Not published"),
 "CIS 2027":     ("Official site", "SCRS series; CIS 2021–2025 in Springer LNNS.", "Not published"),
 "ICST 2027":    ("Listing site", "New series; no past edition found.", "Not published"),
 "ICSISCET 2027":("Official site", "SCRS series; prior editions in Springer LNNS.", "Not published"),
 "ICRTC 2027":   ("Official site", "Past publisher not verified.", "Not published"),
 "ISC 2027":     ("Official site", "Past publisher not verified.", "Not published"),
 "ICAN 2027":    ("Official site", "Past editions' publisher not verified; 2027 partner 'to be updated'.", "Not published"),
 "ICETCI 2027":  ("Official site", "Organizer says 2025 and 2026 proceedings were with Springer; not independently verified.", "Not published"),
 "ICAISS 2027":  ("Official site", "2025 (3rd) proceedings published by IEEE (3 vols).", "Large 3-volume proceedings; not selective."),
 "CVMI 2027":    ("Official site", "IEEE Xplore 2022–2025 (IIIT Allahabad, NIT Rourkela).", "Not published; single-track, moderate volume."),
 "CVIP 2027":    ("Estimated", "Springer CCIS/LNCS every year since 2016.", "Historically ~30–35% (organizer statements in past calls)."),
 "PReMI 2027":   ("Estimated", "Springer LNCS since 2005; 2025: 154 papers in 2 volumes.", "Selective; LNCS-level review."),
 "INDICON 2027": ("Estimated", "IEEE Xplore every year.", "Typically ~30–40%."),
 "ICCCNT 2027":  ("Estimated", "IEEE Xplore every year since 2010.", "Not published; very large volume."),
 "ICARC 2027":   ("Official site", "IEEE Xplore 2021–2026 (IEEE Sri Lanka Section).", "Not published"),
}
TOPIC_RE = __import__("re").compile(r"AI|ML|DL|CV|NLP|vision|learning|intelligen|data science|pattern", __import__("re").I)
def fit_score(r):
    s = 0
    if TOPIC_RE.search(r["Topic area"]): s += 1
    if __import__("re").search(r"power|wireless|broad|business|sustainab", r["Topic area"], __import__("re").I): s -= 1
    if r["Mode"] == "Hybrid": s += 1
    pub = r["Publisher / series"].lower()
    if not any(w in pub for w in ("pending", "not confirmed", "not stated", "not announced", "planned", "proposal", "unconfirmed", "tbc")): s += 1
    if r["Deadline status"].startswith("Confirmed"): s += 1
    if r["Verification"] == "Official site" and "not verified" not in r["Track record"].lower() and "no past" not in r["Track record"].lower(): s += 1
    return max(0, s)
for r in CONFS:
    if r["Acronym"] in QUALITY:
        r["Verification"], r["Track record"], r["Acceptance signals"] = QUALITY[r["Acronym"]]
    elif r["Deadline status"].startswith("Expected"):
        r["Verification"] = "Estimated"
    r["Fit score"] = fit_score(r)

# One entry per refresh; newest first. The refresh task prepends to this list.
CHANGELOG = [
 ("2026-11-17", "Added fit score, verification tier, track record and acceptance columns; closing-soon section and change log on the page. Added ICARC 2027 (Sri Lanka, hybrid IEEE). Coverage now includes hybrid venues outside India."),
 ("2026-11-17", "Window widened to Jan–Aug 2027. Added CIS, ICST, ICSISCET, ICRTC, ISC (moved in), ICISS, ICAISS, CICN; Premier group (CVMI, CVIP, PReMI, INDICON, ICCCNT); COEP CON and IEEE-GCON to watchlist."),
 ("2026-11-16", "First refresh run: fees collected; CONECCT moved to watchlist (deadline 15 Dec 2026); ISC rescheduled to Sep; CVR Puducherry entry replaced by CIMA; added BIDA, ICIVC and several watchlist venues."),
 ("2026-11-16", "Database created: 13 venues, Jan–Mar 2027 window."),
]

wb = Workbook(); wb.remove(wb.active)
sheets = {"IEEE": "IEEE", "Springer": "Springer", "Premier": "Premier", "Publisher TBC": "Publisher TBC",
          "Hybrid only": None, "Watchlist": "Watchlist"}
for title, grp in sheets.items():
    ws = wb.create_sheet(title)
    rows = [r for r in CONFS if (r["Group"] == grp if grp else r["Mode"] == "Hybrid" and r["Group"] != "Watchlist")]
    rows.sort(key=lambda r: r["Submission deadline"])
    ws.append(COLS)
    for r in rows:
        ws.append([r[k] for k in COLS])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill("solid", fgColor="1F3A5F")
    for i, k in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = 40 if k in ("Full name", "Notes", "Website", "Track record", "Acceptance signals") else 18
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "C2"
    if grp:
        with open(f"conferences_2027_{title.replace(' ', '_').lower()}.csv", "w", newline="") as f:
            w = csv.DictWriter(f, COLS); w.writeheader(); w.writerows(rows)
wb.save("conferences_2027.xlsx")
json.dump(CONFS, open("data.json", "w"), indent=1, ensure_ascii=False)
json.dump(CHANGELOG, open("changelog.json", "w"), indent=1, ensure_ascii=False)
print(len(CONFS), "rows written")
html = open("template.html").read().replace("__DATA__", json.dumps(CONFS, ensure_ascii=False)).replace("__CHECKED__", CHECKED).replace("__CHANGELOG__", json.dumps(CHANGELOG, ensure_ascii=False))
open("conferences_2027.html", "w").write(html)
