"""Content Plan slides (Month 1, 2, 3)."""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pptx_builder.core import inches, SLIDE_WIDTH, SLIDE_HEIGHT
from pptx_builder.slides import (
    _textbox, _rect, _rounded_rect, wrap_slide,
    RED, NAVY, DARK, WHITE, LIGHT_BG, GREY, TEXT, SUBTEXT
)


def slide_content_m1():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "06 | CONTENT PLAN - MONTH 1 (Foundation)", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _textbox(4, "f", inches(0.5), inches(1.15), inches(11), inches(0.4), [
        {"text": "THEME: Build Foundation & Brand Awareness | 3-4 posts/week", "size": 1150, "bold": True, "color": RED},
    ])
    s += _rounded_rect(5, "w1", inches(0.4), inches(1.7), inches(5.6), inches(2.4), LIGHT_BG, [
        {"text": "WEEK 1-2: BRAND STORY", "size": 1100, "bold": True, "color": NAVY},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Mon: Meet the Team (carousel)", "size": 950, "color": TEXT},
        {"text": "Wed: Our Manufacturing Process (Reel)", "size": 950, "color": TEXT},
        {"text": "Fri: Before/After Project #1 (carousel)", "size": 950, "color": TEXT},
        {"text": "Sat: Customer Testimonial (Story)", "size": 950, "color": TEXT},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Blog: 'Why Choose a Local Manufacturer'", "size": 950, "color": "2980B9"},
    ])
    s += _rounded_rect(6, "w2", inches(0.4), inches(4.3), inches(5.6), inches(2.4), LIGHT_BG, [
        {"text": "WEEK 3-4: EDUCATION & TRUST", "size": 1100, "bold": True, "color": NAVY},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Mon: Energy Efficiency Tips (infographic)", "size": 950, "color": TEXT},
        {"text": "Wed: Window Style Guide (carousel)", "size": 950, "color": TEXT},
        {"text": "Fri: Before/After Project #2 (Reel)", "size": 950, "color": TEXT},
        {"text": "Sat: FAQ Answer (Story poll)", "size": 950, "color": TEXT},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Blog: 'Calgary Winter Window Prep Guide'", "size": 950, "color": "2980B9"},
    ])
    s += _rounded_rect(7, "pil", inches(6.2), inches(1.7), inches(5.4), inches(5.0), NAVY, [
        {"text": "CONTENT PILLARS", "size": 1200, "bold": True, "color": RED},
        {"text": "", "size": 250, "color": WHITE},
        {"text": "1. Project Showcases (30%)", "size": 1050, "color": WHITE},
        {"text": "   Before/after, time-lapses", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "2. Educational (25%)", "size": 1050, "color": WHITE},
        {"text": "   Tips, guides, how-tos", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "3. Behind the Scenes (20%)", "size": 1050, "color": WHITE},
        {"text": "   Manufacturing, team, process", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "4. Social Proof (15%)", "size": 1050, "color": WHITE},
        {"text": "   Reviews, testimonials, awards", "size": 900, "color": "BDC3C7"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "5. Community (10%)", "size": 1050, "color": WHITE},
        {"text": "   Local events, partnerships", "size": 900, "color": "BDC3C7"},
    ])
    return wrap_slide(s)


def slide_content_m2():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "06 | CONTENT PLAN - MONTH 2 (Engagement)", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _textbox(4, "f", inches(0.5), inches(1.15), inches(11), inches(0.4), [
        {"text": "THEME: Drive Engagement & Community | 4-5 posts/week", "size": 1150, "bold": True, "color": RED},
    ])
    s += _rounded_rect(5, "w5", inches(0.4), inches(1.7), inches(5.6), inches(2.4), LIGHT_BG, [
        {"text": "WEEK 5-6: ENGAGEMENT BOOST", "size": 1100, "bold": True, "color": NAVY},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Mon: 'Guess the Window Style' (poll)", "size": 950, "color": TEXT},
        {"text": "Tue: Installation Day Reel (timelapse)", "size": 950, "color": TEXT},
        {"text": "Thu: Homeowner Spotlight (interview)", "size": 950, "color": TEXT},
        {"text": "Fri: Product Close-up (carousel)", "size": 950, "color": TEXT},
        {"text": "Sat: Weekend DIY tip (Story)", "size": 950, "color": TEXT},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Blog: 'Triple vs Double Pane: Calgary Guide'", "size": 950, "color": "2980B9"},
    ])
    s += _rounded_rect(6, "w7", inches(0.4), inches(4.3), inches(5.6), inches(2.4), LIGHT_BG, [
        {"text": "WEEK 7-8: SEASONAL & LOCAL", "size": 1100, "bold": True, "color": NAVY},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Mon: 'Calgary Weather vs Our Windows' (Reel)", "size": 950, "color": TEXT},
        {"text": "Wed: Local Builder Partnership Feature", "size": 950, "color": TEXT},
        {"text": "Thu: Energy Savings Calculator Post", "size": 950, "color": TEXT},
        {"text": "Fri: Project Reveal (Reel + carousel)", "size": 950, "color": TEXT},
        {"text": "Sat: Customer review highlight", "size": 950, "color": TEXT},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Blog: '5 Signs You Need Window Replacement'", "size": 950, "color": "2980B9"},
    ])
    s += _rounded_rect(7, "cr", inches(6.2), inches(1.7), inches(5.4), inches(5.0), "5B2D8E", [
        {"text": "CREATIVE CONTENT IDEAS", "size": 1200, "bold": True, "color": WHITE},
        {"text": "", "size": 250, "color": WHITE},
        {"text": "'Window Transformation Tuesday'", "size": 1050, "color": WHITE},
        {"text": "   Weekly series - project reveals", "size": 900, "color": "E8D5F5"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "'Ask Apollo' Q&A Stories", "size": 1050, "color": WHITE},
        {"text": "   Answer homeowner questions", "size": 900, "color": "E8D5F5"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "'Made in Calgary' Series", "size": 1050, "color": WHITE},
        {"text": "   Show local manufacturing pride", "size": 900, "color": "E8D5F5"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "'Energy Bill Challenge'", "size": 1050, "color": WHITE},
        {"text": "   Before/after energy savings", "size": 900, "color": "E8D5F5"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "'Meet Your Installer' Reels", "size": 1050, "color": WHITE},
        {"text": "   Humanize the brand", "size": 900, "color": "E8D5F5"},
    ])
    return wrap_slide(s)


def slide_content_m3():
    s = ""
    s += _rect(2, "h", 0, 0, SLIDE_WIDTH, inches(1.1), NAVY)
    s += _textbox(3, "ht", inches(0.5), inches(0.15), inches(10), inches(0.9), [
        {"text": "06 | CONTENT PLAN - MONTH 3 (Conversion)", "size": 1900, "bold": True, "color": WHITE},
    ])
    s += _textbox(4, "f", inches(0.5), inches(1.15), inches(11), inches(0.4), [
        {"text": "THEME: Convert Followers to Leads | 4-5 posts/week + daily Stories", "size": 1150, "bold": True, "color": RED},
    ])
    s += _rounded_rect(5, "w9", inches(0.4), inches(1.7), inches(5.6), inches(2.4), LIGHT_BG, [
        {"text": "WEEK 9-10: SOCIAL PROOF & OFFERS", "size": 1100, "bold": True, "color": NAVY},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Mon: Customer Story Video (Reel)", "size": 950, "color": TEXT},
        {"text": "Tue: Limited-Time Offer Graphic", "size": 950, "color": TEXT},
        {"text": "Thu: '25 Years of Craftsmanship' carousel", "size": 950, "color": TEXT},
        {"text": "Fri: Side-by-side comparison post", "size": 950, "color": TEXT},
        {"text": "Sat: Free consultation CTA Story", "size": 950, "color": TEXT},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Blog: 'How to Choose the Right Windows'", "size": 950, "color": "2980B9"},
    ])
    s += _rounded_rect(6, "w11", inches(0.4), inches(4.3), inches(5.6), inches(2.4), LIGHT_BG, [
        {"text": "WEEK 11-12: LEAD GENERATION", "size": 1100, "bold": True, "color": NAVY},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Mon: Project Portfolio Reel (best work)", "size": 950, "color": TEXT},
        {"text": "Wed: 'Why Apollo' comparison post", "size": 950, "color": TEXT},
        {"text": "Thu: Financing options infographic", "size": 950, "color": TEXT},
        {"text": "Fri: End-of-quarter celebration", "size": 950, "color": TEXT},
        {"text": "Sat: 'Book Your Free Quote' CTA", "size": 950, "color": TEXT},
        {"text": "", "size": 150, "color": TEXT},
        {"text": "Blog: 'Window ROI for Calgary Homeowners'", "size": 950, "color": "2980B9"},
    ])
    s += _rounded_rect(7, "conv", inches(6.2), inches(1.7), inches(5.4), inches(5.0), RED, [
        {"text": "CONVERSION TACTICS", "size": 1200, "bold": True, "color": WHITE},
        {"text": "", "size": 250, "color": WHITE},
        {"text": "Lead Magnets:", "size": 1050, "bold": True, "color": WHITE},
        {"text": "Free Window Style Guide PDF", "size": 950, "color": "FFDDD5"},
        {"text": "Energy Savings Calculator", "size": 950, "color": "FFDDD5"},
        {"text": "Free In-Home Consultation", "size": 950, "color": "FFDDD5"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "CTA Strategies:", "size": 1050, "bold": True, "color": WHITE},
        {"text": "Story swipe-up to quote form", "size": 950, "color": "FFDDD5"},
        {"text": "Bio link to booking page", "size": 950, "color": "FFDDD5"},
        {"text": "DM automation for inquiries", "size": 950, "color": "FFDDD5"},
        {"text": "", "size": 150, "color": WHITE},
        {"text": "Retargeting:", "size": 1050, "bold": True, "color": WHITE},
        {"text": "Website visitors > social ads", "size": 950, "color": "FFDDD5"},
        {"text": "Engaged audience > offers", "size": 950, "color": "FFDDD5"},
    ])
    return wrap_slide(s)
