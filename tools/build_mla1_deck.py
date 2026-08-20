# -*- coding: utf-8 -*-
"""MLA Level 1 — fundamentals, in the KLC-SM house style.

Built from the two sets of Lube_Win_MLA_1 notes and the screen captures
embedded in them. Nine slides, running from maintenance strategy through to
oil analysis. Body text and speaker notes are English so the deck can be
reviewed with an English-speaking manager; the captures are reused as-is.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
import deck_style as ds
from deck_style import (L, R, COLW, TOP, INK, GREY, ACCENT, NAVY, RED, WHITE,
                        LIGHT, chip, textbox, para, rect, table, callout,
                        blank_slide, notes)

ds.LATIN_FONT = "Calibri"

SCRATCH = "/tmp/claude-0/-home-user-github-test/eae68b7d-8c95-56e3-85c5-ff5decc49086/scratchpad/"
MEDIA = SCRATCH + "mla1/"    # captures from the first-half notes
MEDIA2 = SCRATCH + "mla/"    # captures from the second-half notes
OUT = "/home/user/github-test/MLA1_Fundamentals.pptx"

prs = Presentation()
prs.slide_width, prs.slide_height = ds.SW, ds.SH

# ---------------------------------------------------------------- slide 1
s = blank_slide(prs, "Lubrication Strategy Follows Criticality")

y = chip(s, L, TOP, COLW, "Why lubrication is preventive maintenance")
tb, tf = textbox(s, L, y, COLW, Inches(1.5))
para(tf, "If roughly 80% of failures trace back to contaminated or degraded oil, "
         "then every lubrication action counts as preventive maintenance — not housekeeping.",
     13, False, INK, 8, first=True, line=1.25)
para(tf, "The same logic as heart health: diet and a blood test every six months "
         "cost very little. Surgery does not.", 13, False, GREY, 0, line=1.25)

y2 = y + Inches(1.55)
y2 = chip(s, L, y2, COLW, "The six functions of a lubricant")
table(s, L, y2, COLW,
      [[("Reduce friction", False, INK, PP_ALIGN.CENTER), ("Reduce wear", False, INK, PP_ALIGN.CENTER)],
       [("Prevent corrosion", False, INK, PP_ALIGN.CENTER), ("Remove heat", False, INK, PP_ALIGN.CENTER)],
       [("Remove contaminants", False, INK, PP_ALIGN.CENTER), ("Transmit fluid power", False, INK, PP_ALIGN.CENTER)]],
      [Inches(2.975), Inches(2.975)], head=False, row_h=Inches(0.42), size=12.5)

y = chip(s, R, TOP, COLW, "There is no single right answer — spend where it counts")
end = table(s, R, y, COLW,
            [["Equipment character", "What to specify"],
             [("Cheap, easy to replace,\nnot critical", False, INK, PP_ALIGN.LEFT),
              ("A basic lubricant is enough", False, INK, PP_ALIGN.LEFT)],
             [("Spare held on board,\nquick swap", False, INK, PP_ALIGN.LEFT),
              ("Standard grade", False, INK, PP_ALIGN.LEFT)],
             [("Stern tube, thruster —\ncostly, hard to reach,\nexpensive downtime", True, ACCENT, PP_ALIGN.LEFT),
              ("Specify the best available", True, ACCENT, PP_ALIGN.LEFT)]],
            [Inches(2.85), Inches(3.10)], row_h=Inches(0.62), size=12)

callout(s, R, end + Inches(0.25), COLW,
        "Criticality — not price — sets the specification.", h=Inches(0.7),
        size=14, bold=True, color=NAVY)

# ---------------------------------------------------------------- slide 2
s = blank_slide(prs, "Friction and the Four Lubrication Regimes")

y = chip(s, L, TOP, COLW, "Stribeck curve")
s.shapes.add_picture(MEDIA + "image5.png", L, y, width=COLW)   # 1264x634 -> 2.98"
tb, tf = textbox(s, L, y + Inches(3.12), COLW, Inches(0.32))
para(tf, "Friction against ZN/P, and where each regime shows up on board",
     12, False, GREY, 0, first=True, align=PP_ALIGN.CENTER)

callout(s, L, y + Inches(3.60), COLW,
        "Slow and heavily loaded \u2192 higher viscosity, so the film forms sooner.",
        h=Inches(0.7), size=14, bold=True, color=NAVY)

y = chip(s, R, TOP, COLW, "The intuition")
s.shapes.add_picture(MEDIA + "image3.png", R, y, width=COLW)   # 1386x673 -> 2.89"
tb, tf = textbox(s, R, y + Inches(3.02), COLW, Inches(1.5))
para(tf, "Push a block across concrete and friction is highest before it moves — "
         "and barely depends on speed.", 13, False, INK, 8, first=True, line=1.25)
para(tf, "Add a lubricant and friction drops sharply. Keep raising the speed and it "
         "climbs again, the way a ship has to push aside more water the faster it goes: "
         "now the oil itself is what must be sheared.", 13, False, INK, 0, line=1.25)


# ---------------------------------------------------------------- slide 3
s = blank_slide(prs, "Base Oils — Mineral and Synthetic")

y = chip(s, L, TOP, COLW, "Two ways to make a base oil")
s.shapes.add_picture(MEDIA + "image9_crop.png", L, y, width=COLW)  # 1320x583 -> 2.63"
tb, tf = textbox(s, L, y + Inches(2.78), COLW, Inches(1.5))
para(tf, "Mineral — refined from crude oil. A box of mixed Lego: many shapes, "
         "many sizes, whatever the barrel gave you.", 13, False, INK, 8, first=True, line=1.25)
para(tf, "Synthetic — known molecules assembled into the structure you actually want.",
     13, False, INK, 0, line=1.25)

y = chip(s, R, TOP, COLW, "API base oil groups")
s.shapes.add_picture(MEDIA + "image10.png", R, y, width=COLW)  # 1301x632 -> 2.89"

yy = y + Inches(3.05)
tb, tf = textbox(s, R, yy, COLW, Inches(1.6))
para(tf, "Group III is sold as synthetic in most markets, but it is still refined "
         "mineral stock — synthetic in name.", 12.5, False, INK, 6, first=True, line=1.2)
para(tf, "Group IV (PAO) is the workhorse synthetic — gears and some compressors.",
     12.5, False, INK, 6, line=1.2)
para(tf, "Group V is everything that fits nowhere else, which does not make it better.",
     12.5, False, INK, 0, line=1.2)

callout(s, R, yy + Inches(1.20), COLW,
        "Viscosity Index = how much the oil thins as it heats. Higher VI, flatter response.",
        h=Inches(0.65), size=13, bold=True, color=NAVY)

# ---------------------------------------------------------------- slide 4
s = blank_slide(prs, "Additives, Viscosity and Grease")

y = chip(s, L, TOP, COLW, "Additive families")
s.shapes.add_picture(MEDIA + "image11_crop.png", L + Inches(0.62), y,
                     width=Inches(4.70))                  # 1009x641 -> 2.99"

y2 = y + Inches(3.12)
tb, tf = textbox(s, L, y2, COLW, Inches(0.55))
para(tf, "Additives are the most expensive part of the blend — do not pay for what "
         "the application does not need.", 12.5, True, ACCENT, 0, first=True, line=1.2)

table(s, L, y2 + Inches(0.58), COLW,
      [[("Open to atmosphere", False, INK, PP_ALIGN.LEFT), ("More oxidation inhibitor", False, INK, PP_ALIGN.LEFT)],
       [("Heavily loaded", False, INK, PP_ALIGN.LEFT), ("Extreme pressure (EP)", False, INK, PP_ALIGN.LEFT)],
       [("Hydraulic pump", False, INK, PP_ALIGN.LEFT), ("Antiwear (AW)", False, INK, PP_ALIGN.LEFT)],
       [("Compressor, turbine", False, INK, PP_ALIGN.LEFT), ("R&O — rust and oxidation", False, INK, PP_ALIGN.LEFT)]],
      [Inches(2.45), Inches(3.50)], head=False, row_h=Inches(0.32), size=12)

y = chip(s, R, TOP, COLW, "The three properties that matter most")
end = table(s, R, y, COLW,
            [[("1", True, NAVY, PP_ALIGN.CENTER), ("Viscosity", False, INK, PP_ALIGN.LEFT)],
             [("2", True, NAVY, PP_ALIGN.CENTER), ("Viscosity against temperature", False, INK, PP_ALIGN.LEFT)],
             [("3", True, NAVY, PP_ALIGN.CENTER), ("Viscosity against pressure", False, INK, PP_ALIGN.LEFT)]],
            [Inches(0.65), Inches(5.30)], head=False, row_h=Inches(0.36), size=12.5)
tb, tf = textbox(s, R, end + Inches(0.12), COLW, Inches(0.5))
para(tf, "Kinematic viscosity is the figure normally quoted; ISO VG is the grading standard.",
     12.5, False, GREY, 0, first=True, line=1.2)

y3 = end + Inches(0.75)
y3 = chip(s, R, y3, COLW, "Grease = base oil + thickener + additives")
tb, tf = textbox(s, R, y3, COLW, Inches(1.6))
para(tf, "Think of a sponge holding water. Squeeze it and water comes out; release it "
         "and the sponge takes the water back.", 13, False, INK, 8, first=True, line=1.25)
para(tf, "The thickener matrix holds the base oil until the bearing turns and load comes "
         "on. The oil is pushed out, and it is the base oil that does the lubricating.",
     13, False, INK, 0, line=1.25)

callout(s, R, y3 + Inches(1.75), COLW,
        "Consistency is graded by worked penetration (NLGI). NLGI 2 is the general "
        "multipurpose grade.", h=Inches(0.8), size=13, bold=True, color=NAVY)




def crop_bottom(name, keep):
    """Trim the burned-in Korean subtitle strip off a screen capture."""
    from PIL import Image
    src = MEDIA2 + name
    dst = SCRATCH + "trim_" + name
    im = Image.open(src)
    w, h = im.size
    im.crop((0, 0, w, int(h * keep))).save(dst)
    return dst


# ---------------------------------------------------------------- slide 5
s = blank_slide(prs, "Selecting for the Application")

y = chip(s, L, TOP, COLW, "Rolling bearings — viscosity from size and speed")
s.shapes.add_picture(crop_bottom("image4.png", 0.94), L, y, width=COLW)   # 989x587 -> 3.53"
tb, tf = textbox(s, L, y + Inches(3.66), COLW, Inches(1.4))
para(tf, "Read the viscosity the bearing needs from its mean diameter and speed, "
         "then read what the oil actually gives at working temperature.",
     12.5, False, INK, 8, first=True, line=1.25)
para(tf, "κ = actual ÷ rated. Below 1 the film is too thin.",
     13, True, ACCENT, 0, line=1.2)

y = chip(s, R, TOP, COLW, "Hydraulics — zinc and viscosity index")
s.shapes.add_picture(MEDIA2 + "image8.png", R + Inches(0.73), y,
                     width=Inches(4.50))                        # 1317x912 -> 3.11"
yy = y + Inches(3.25)
tb, tf = textbox(s, R, yy, COLW, Inches(0.95))
para(tf, "Zinc-free where toxicity matters or servo valves are fitted; high VI for "
         "wide temperature swings. Gear oils are specified by API GL class — "
         "GL-4 for manual gearboxes, GL-5 for hypoid axles.",
     12.5, False, INK, 0, first=True, line=1.25)
callout(s, R, yy + Inches(1.05), COLW,
        "Cost rises toward the top right. Expensive is not automatically correct.",
        h=Inches(0.62), size=13, bold=True, color=NAVY)

notes(s, "Now, how do we pick the oil for a real machine.\n\n"
         "For rolling bearings we compare two numbers. First, the viscosity the "
         "bearing needs — that comes from its size and its speed. Second, the "
         "viscosity the oil really has at working temperature. Divide one by the "
         "other and you get kappa. Above one is good. Below one means the film is "
         "too thin and the metal starts to touch.\n\n"
         "For hydraulics there are only two questions. Zinc or zinc free, and "
         "standard or high viscosity index. Zinc free is for low toxicity, and for "
         "servo valves, because zinc can break down and leave deposits in very "
         "small clearances. High VI is for wide temperature swings and mobile "
         "equipment.\n\n"
         "Cost goes up as you move to the top right of that box. But expensive is "
         "not automatically the right answer — come back to slide one, it "
         "depends on the machine.\n\n"
         "For gears, check the API GL class on the spec. GL-4 for manual gearboxes, "
         "GL-5 for hypoid axles.")

# ---------------------------------------------------------------- slide 6
s = blank_slide(prs, "Contamination Control I — Particles")

y = chip(s, L, TOP, COLW, "How to read an ISO cleanliness code")
s.shapes.add_picture(crop_bottom("image16.png", 0.92), L, y, width=COLW)  # 1991x1113 -> 3.33"
tb, tf = textbox(s, L, y + Inches(3.46), COLW, Inches(1.4))
para(tf, "Three numbers, for particles larger than 4, 6 and 14 µm. Each step "
         "doubles the count — so 20 to 22 is four times the dirt.",
     12.5, False, INK, 8, first=True, line=1.25)
para(tf, "Counted by microscopy (ISO 4407), light extinction (ISO 11500), patch "
         "test, pore blockage or laser net fines.", 12.5, False, GREY, 0, line=1.25)

y = chip(s, R, TOP, COLW, "Filters — and when they stop filtering")
s.shapes.add_picture(crop_bottom("image18.png", 0.82), R, y, width=COLW)  # 1901x1207 -> 3.78"
tb, tf = textbox(s, R, y + Inches(3.90), COLW, Inches(1.0))
para(tf, "Surface media stop anything wider than the gap; depth media trap fine "
         "particles inside the fibres; magnets catch large steel.",
     12.5, False, INK, 8, first=True, line=1.25)
para(tf, "ΔP shows how blocked the element is. Past a set point the bypass "
         "opens and unfiltered oil goes straight through.",
     12.5, True, RED, 0, line=1.25)

notes(s, "Around eighty percent of hydraulic failures start with dirty oil, so we "
         "measure cleanliness — we do not guess it.\n\n"
         "The ISO code is three numbers. Particles bigger than four microns, bigger "
         "than six, and bigger than fourteen. Each step up the scale is double. So "
         "going from twenty to twenty-two is not a little worse, it is four times "
         "more dirt. Walk through the table on the slide once, slowly.\n\n"
         "There are several ways to count. A microscope, a laser, a patch test, or "
         "a blocked-pore method. The patch test also shows you what the particles "
         "look like, which tells you where they came from.\n\n"
         "Now filters. A surface filter is a screen — nothing wider than the "
         "gap gets through. A depth filter is a mat of fibres and catches small "
         "particles inside it. Magnetic plugs pick up large steel pieces.\n\n"
         "Last point, and this is the one people forget. Delta P tells you how "
         "blocked the filter is. When it gets too high the bypass valve opens, and "
         "from that moment the oil is not being filtered at all.")

# ---------------------------------------------------------------- slide 7
s = blank_slide(prs, "Contamination Control II — Water and Air")

y = chip(s, L, TOP, COLW, "Water attacks the additives, not just the metal")
s.shapes.add_picture(crop_bottom("image12.png", 0.92), L, y, width=COLW)  # 1982x1080 -> 3.24"
tb, tf = textbox(s, L, y + Inches(3.36), COLW, Inches(1.5))
para(tf, "Oils are hygroscopic — they pull water in, even out of the air. "
         "Water appears dissolved, emulsified, then free.",
     12.5, False, INK, 8, first=True, line=1.25)
para(tf, "Free water drains. Dissolved and emulsified are the hard ones — and "
         "they drive the additive reactions backwards, making acid.",
     12.5, False, INK, 0, line=1.25)

y = chip(s, R, TOP, COLW, "Air — foam is not the dangerous one")
s.shapes.add_picture(MEDIA2 + "image15.png", R, y, width=COLW)  # 1975x1162 -> 3.50"
tb, tf = textbox(s, R, y + Inches(3.62), COLW, Inches(1.5))
para(tf, "Free air sits on top as foam. Dissolved air is invisible. Entrained air "
         "— bubbles carried inside the oil — is the damaging one.",
     12.5, False, INK, 8, first=True, line=1.25)
para(tf, "It reaches the pump and causes cavitation, spongy control and low oil "
         "pressure trips.", 12.5, True, RED, 0, line=1.25)

notes(s, "Water first.\n\n"
         "Oil pulls water in. Many oils are hygroscopic, which means they take "
         "moisture straight out of the air. Water shows up three ways. Dissolved, "
         "which you cannot see at all. Emulsified, which makes the oil look cloudy. "
         "And free water, which settles at the bottom.\n\n"
         "Free water is the easy one — you drain it. Dissolved and emulsified "
         "are the hard ones to remove.\n\n"
         "And water does more than rust metal. Look at the reaction on the slide. "
         "The additives were built by taking water out. Put water back in and the "
         "reaction runs backwards. The additive breaks apart and you get acid, so "
         "TAN goes up.\n\n"
         "Now air. Everybody worries about the foam they can see on top of the tank. "
         "That is not the dangerous one. The dangerous one is entrained air — "
         "small bubbles carried inside the oil. Those go into the pump and cause "
         "cavitation, vibration, spongy controls, and low oil pressure alarms.\n\n"
         "One more thing worth saying: most foam problems are not the oil's fault. "
         "They come from contamination or from oxidation.")

# ---------------------------------------------------------------- slide 8
s = blank_slide(prs, "Degradation and Wear")

y = chip(s, L, TOP, COLW, "Three ways the oil itself goes bad")
s.shapes.add_picture(crop_bottom("image11.png", 0.93), L + Inches(0.17), y,
                     width=Inches(5.60))                        # 1993x1255 -> 3.53"
callout(s, L, y + Inches(3.66), COLW,
        "Above 60 °C, every extra 10 °C halves the oil's service life.",
        h=Inches(0.62), size=13, bold=True, color=NAVY)
tb, tf = textbox(s, L, y + Inches(4.42), COLW, Inches(0.5))
para(tf, "Additives leave three ways too: they decompose, they are separated out by "
         "filters and centrifuges, or they adsorb onto surfaces and particles.",
     12, False, GREY, 0, first=True, line=1.2)

y = chip(s, R, TOP, COLW, "Four ways surfaces wear")
s.shapes.add_picture(MEDIA2 + "image19.png", R, y, width=COLW)  # 1872x1010 -> 3.21"
end = table(s, R, y + Inches(3.34), COLW,
            [[("Adhesive", True, INK, PP_ALIGN.LEFT), ("Metal welds and tears — scuffing", False, INK, PP_ALIGN.LEFT)],
             [("Abrasive", True, INK, PP_ALIGN.LEFT), ("Hard particles cut — scoring", False, INK, PP_ALIGN.LEFT)],
             [("Corrosive", True, INK, PP_ALIGN.LEFT), ("Acid and water attack the surface", False, INK, PP_ALIGN.LEFT)],
             [("Surface fatigue", True, INK, PP_ALIGN.LEFT), ("Repeated load — micropitting", False, INK, PP_ALIGN.LEFT)]],
            [Inches(1.75), Inches(4.20)], head=False, row_h=Inches(0.34), size=12)

notes(s, "Three ways the oil itself goes bad.\n\n"
         "Oxidation. Oxygen attacks the oil, makes acid, thickens it and darkens it. "
         "Heat, metals, contamination and water all speed it up.\n\n"
         "Nitration. This comes from combustion gases, so it is an engine problem. "
         "It also makes acid, and it leaves a red varnish.\n\n"
         "Thermal degradation, above about two hundred and twenty degrees. This one "
         "can go both ways — coke and carbon make the oil thicker, but cracking "
         "the molecules makes it thinner.\n\n"
         "If the crew remembers one number, make it this one: above sixty degrees, "
         "every extra ten degrees cuts the oil life in half.\n\n"
         "Additives disappear as well. They break down, or the filter and the "
         "centrifuge take them out, or they stick to surfaces and to particles.\n\n"
         "On the right, four wear types. Adhesive is metal welding to metal and "
         "tearing — scuffing. Abrasive is hard particles cutting grooves — "
         "scoring. Corrosive is acid and water. Surface fatigue is repeated loading, "
         "and it shows as micropitting.")

# ---------------------------------------------------------------- slide 9
s = blank_slide(prs, "Oil Analysis and the Programme")

y = chip(s, L, TOP, COLW, "ICP metals — where each element points")
s.shapes.add_picture(MEDIA2 + "image14.png", L, y, width=COLW)  # 1926x1186 -> 3.66"
tb, tf = textbox(s, L, y + Inches(3.78), COLW, Inches(1.3))
para(tf, "Sampling decides the result. Use a clean bottle, and take the sample where "
         "the answer is — after the pump and before the filter when hunting "
         "pump wear.", 12.5, False, INK, 8, first=True, line=1.25)
para(tf, "Ca, Mg, Zn and P are usually additives, not wear.",
     12.5, True, ACCENT, 0, line=1.2)

y = chip(s, R, TOP, COLW, "Relubrication — little and often")
s.shapes.add_picture(crop_bottom("image22.png", 0.78), R + Inches(0.48), y,
                     width=Inches(5.00))                        # 1100x572 -> 2.60"
tb, tf = textbox(s, R, y + Inches(2.74), COLW, Inches(1.2))
para(tf, "Manual greasing swings between too much and too little. Small, frequent "
         "doses hold the safe band — which is what automatic lubricators do "
         "well, provided somebody still checks them.",
     12.5, False, INK, 0, first=True, line=1.25)

callout(s, R, y + Inches(3.90), COLW,
        "Keep it clean, keep it cool, keep it dry.", h=Inches(0.75),
        size=16, bold=True, color=NAVY)
tb, tf = textbox(s, R, y + Inches(4.72), COLW, Inches(0.32))
para(tf, "Check the delivery against the order, and do not store drums hot.",
     12, False, GREY, 0, first=True, line=1.2)

notes(s, "Oil analysis only works if the sample represents the machine. Where and "
         "how you take it matters more than which laboratory you use.\n\n"
         "Use a clean bottle — a dirty bottle gives you a dirty result and you "
         "will chase a problem that is not there. Take the sample where the answer "
         "is. If you are looking for pump wear, sample after the pump and before the "
         "filter, because the filter removes the evidence.\n\n"
         "ICP tells you which metals are in the oil, and each one points somewhere. "
         "Iron is engine parts or rust. Copper is bearings, bushings and oil coolers. "
         "Silicon is dirt. Sodium and potassium usually mean coolant or seawater.\n\n"
         "Be careful with calcium, magnesium, zinc and phosphorus. Those are normally "
         "additives, not wear metals. Do not raise an alarm for those.\n\n"
         "On the right, greasing. Greasing by hand swings between too much and too "
         "little. Small amounts, more often, keeps it inside the safe band. That is "
         "what automatic lubricators are good at — but somebody still has to "
         "check them. Fit and forget is how you find an empty unit two years later.\n\n"
         "Close the session with three words. Clean, cool, dry. And two habits: "
         "check what was delivered against what was ordered, and do not store drums "
         "somewhere hot.")

# ---------------------------------------------------------------- notes 1-4
_slides = list(prs.slides)

notes(_slides[0],
      "Start here. If most failures come from dirty or worn-out oil, then looking "
      "after the oil is real maintenance work — not cleaning.\n\n"
      "Use the health example. Eating well and a blood test twice a year costs very "
      "little. Heart surgery does not. Same idea with machines.\n\n"
      "Then the second point, and it is the important one. We do not buy the best "
      "oil for everything. For a small pump with a spare in the store, a basic oil "
      "is fine. For a stern tube or a thruster — expensive, hard to reach, and "
      "the ship stops if it fails — we buy the best available.\n\n"
      "So the rule is simple: how critical the machine is decides the oil. Not the "
      "price of the drum.\n\n"
      "Finish by reading the six functions. Most people only think of the first two.")

notes(_slides[1],
      "This curve is the core idea of lubrication. Take your time here.\n\n"
      "On the left, metal touches metal. In the middle, the film is starting to "
      "build. On the right, a full oil film carries the load and the surfaces never "
      "touch.\n\n"
      "The axis is Z N over P. Z is viscosity, N is speed, P is load.\n\n"
      "Use the block example. Pushing a heavy box across concrete is hardest right "
      "at the start, and going faster does not help. Now add oil and it becomes much "
      "easier. But keep increasing the speed and it gets harder again — because "
      "now you are pushing the oil itself out of the way, like a ship pushing water.\n\n"
      "The practical point for the crew: a slow machine carrying a heavy load needs "
      "thicker oil, because the film needs help to form. Start-up and shut-down are "
      "always boundary — that is when most wear happens.")

notes(_slides[2],
      "Two ways to make a base oil.\n\n"
      "Mineral oil comes from crude. Think of a box of mixed Lego — many shapes "
      "and sizes, because that is what the barrel gave us.\n\n"
      "Synthetic is built from chosen molecules, so every piece is the same shape. "
      "That is why it behaves more predictably.\n\n"
      "Now the API table. Groups one and two are mineral. Group three is still "
      "refined mineral stock, but most markets are allowed to sell it as synthetic. "
      "Worth knowing when you compare two quotations. Group four, PAO, is the common "
      "true synthetic — gears and some compressors. Group five is everything "
      "else, and that does not mean better.\n\n"
      "Last point. Viscosity index tells you how much the oil thins when it gets "
      "hot. Higher VI, flatter response, fewer surprises between cold start and "
      "full load.")

notes(_slides[3],
      "Additives do three jobs. Change how the oil performs, protect the oil itself, "
      "and protect the metal. Point at each column as you say it.\n\n"
      "Then the money point. Additives are the most expensive part of the blend, so "
      "we only pay for what the machine actually needs. Open to the air — more "
      "antioxidant. Heavy load — EP. Hydraulic pump — antiwear. Compressor "
      "or turbine — R and O, rust and oxidation.\n\n"
      "On the right: if they remember one word from the whole session, it is "
      "viscosity. Viscosity itself, how it changes with temperature, and how it "
      "changes with pressure. Kinematic viscosity is the number on the datasheet, "
      "and ISO VG is the grade.\n\n"
      "Finally grease. Grease is oil held in a sponge. The thickener is the sponge; "
      "the base oil does the actual lubricating. Squeeze it and the oil comes out. "
      "That is exactly what happens when the bearing turns and load comes on.\n\n"
      "NLGI is the consistency grade, measured by worked penetration. NLGI 2 is the "
      "normal all-purpose grease. NLGI 0 is for centralised systems, because it has "
      "to be pumped.")

prs.save(OUT)
print("saved", OUT, "| slides:", len(prs.slides._sldIdLst))
