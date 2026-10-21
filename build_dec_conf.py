#!/usr/bin/env python3
"""
build_dec_conf.py
Generates Cell_Leaders_Conference_Faithfulness_in_Stewardship.docx and .md versions
following the Refined Claude Standard and Saints Community Church Study Group Guidelines.
"""

import os
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_documents(output_docx_paths, output_md_paths):
    doc = Document()

    # Page setup: Letter size, 1.0 inch margins
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Style defaults
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0x11, 0x18, 0x27) # Dark neutral

    md_lines = []

    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(18)
        md_lines.append(f"# {text}")

    def add_subtitle(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(10)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        md_lines.append(f"## *{text}*\n\n---\n")

    def h(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"### {text}\n")
        return p

    def ref(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"***{text}***")
        return p

    def v(text):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"> ***{text}***  ")
        return p

    def sp():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        md_lines.append("")
        return p

    def b(runs_data):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        md_para = ""
        for r_text, r_bold, r_italic in runs_data:
            run = p.add_run(r_text)
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            if r_bold:
                run.bold = True
            if r_italic:
                run.italic = True

            # Markdown conversion
            seg = r_text
            if r_bold and r_italic:
                seg = f"***{seg}***"
            elif r_bold:
                seg = f"**{seg}**"
            elif r_italic:
                seg = f"*{seg}*"
            md_para += seg

        md_lines.append(md_para + "\n")
        return p

    # =========================================================================
    # DOCUMENT CONTENT
    # =========================================================================

    add_title("Cell Leaders Conference")
    add_subtitle("Faithfulness in Stewardship: The Investment of Lives and the Four Demands of Pastoral Oversight")

    # -------------------------------------------------------------------------
    # Section 1
    # -------------------------------------------------------------------------
    h("The Requirement of a Steward: Faithfulness, Loyalty, and Pure Motivation")
    ref("1 Corinthians 4:1-2")
    v("[1] Let a man so account of us, as of the ministers of Christ, and stewards of the mysteries of God.")
    v("[2] Moreover it is required in stewards, that a man be found faithful.")
    sp()

    b([
        ("In ", False, False),
        ("1 Corinthians 4:1-2", True, False),
        (", the Apostle Paul defines the baseline evaluation of ministerial identity. A minister is not accounted by worldly titles, public acclaim, or political influence, but as a servant of Christ and a steward of divine mysteries. The governing requirement laid upon a steward is singularly defined: ", False, False),
        ("moreover, it is required in stewards, that a man be found faithful", True, False),
        (". The Greek term translated 'faithful' is ", False, False),
        ("pistos", True, True),
        (", derived from the verb ", False, False),
        ("pisteuo", True, True),
        (". It denotes absolute trustworthiness, unwavering reliability, and steadfast integrity. When Scripture affirms that God is faithful (", False, False),
        ("1 Corinthians 10:13", True, False),
        ("), it testifies that God is ever reliable, never reneges on His Word, and remains constant regardless of opposing circumstances.", False, False)
    ])
    sp()

    b([
        ("Faithfulness in Christian stewardship is synonymous with ", False, False),
        ("loyalty", True, False),
        (". Loyalty is the unreserved, one-hundred-percent commitment of heart and labor to a divine cause. It demands that a leader remain constant in allegiance to the assignment, the local church, and the spiritual oversight committed to their trust. There is a fundamental, irreconcilable difference between ", False, False),
        ("faithfulness and personal ambition", True, False),
        (". Ambitious individuals attach themselves to a church or ministry to rise through the ranks, seek visibility, or exploit spiritual platforms for personal advancement. The true test that unmasks mere ambition is when personal desires, positional promotions, or recognition are withheld. When an ambitious individual experiences slight alterations, corrections, or perceived delays, they immediately alter their posture, murmur, or withdraw. Conversely, the faithful individual is impervious to fluctuating circumstances because their service is anchored in God, not personal entitlement.", False, False)
    ])
    sp()

    ref("Ruth 1:16")
    v("[16] And Ruth said, Intreat me not to leave thee, or to return from following after thee: for whither thou goest, I will go; and where thou lodgest, I will lodge: thy people shall be my people, and thy God my God.")
    sp()

    b([
        ("Ruth stands as an enduring Old Testament illustration of uncalculating loyalty. When Naomi urged her daughters-in-law to return to their homeland, Orpah kissed her mother-in-law and returned, but Ruth clave unto her. Ruth possessed no earthly promise of marriage, wealth, or prestige in Israel. Her confession was an absolute covenant of dedication: ", False, False),
        ("*'Whither thou goest, I will go; and where thou lodgest, I will lodge: thy people shall be my people, and thy God my God.'*", True, True),
        (" She was not calculating perks or advantages; her devotion was pure, unadulterated faithfulness.", False, False)
    ])
    sp()

    ref("2 Kings 3:11")
    v("[11] Here is Elisha the son of Shaphat, which poured water on the hands of Elijah.")
    sp()

    b([
        ("A parallel pattern of faithful stewardship is embodied in Elisha. Historical evidence indicates that Elisha emerged from a prosperous agricultural family and was likely older, yet he willingly subjected himself to menial, steadfast servanthood under Elijah. When kings sought a prophet of God, the testimony that accredited Elisha was: ", False, False),
        ("*'Here is Elisha the son of Shaphat, which poured water on the hands of Elijah.'*", True, True),
        (" He was known not by oratorical titles, but by dependable, humble service. As Elijah moved through Gilgal, Bethel, Jericho, and toward Jordan (", False, False),
        ("2 Kings 2", True, False),
        ("), the sons of the prophets remained idle observers, questioning why Elisha persisted in following a master who was about to be taken away. Elisha silenced their cynicism, pressed past every barrier, and stood with Elijah to the end.", False, False)
    ])
    sp()

    b([
        ("When Elijah asked what he should do for him, Elisha requested a double portion of his spirit—which was the legal right of the firstborn inheritor to maintain the spiritual continuity of the prophetic office. When Elijah was caught up and the mantle fell, Elisha did not exalt himself. Striking the waters of the Jordan, he cried: ", False, False),
        ("*'Where is the LORD God of Elijah?'*", True, True),
        (" He did not cry, 'Where is the God of Elisha?' He made it known to all that he was merely carrying on Elijah's ministry under the same God. In identical fashion, Joshua honored Moses, never exploiting Moses' passing to establish a divergent, self-aggrandizing reputation, but faithfully executing the instructions given through the servant of the Lord.", False, False)
    ])
    sp()

    b([
        ("This exposes the ultimate litmus test of stewardship: ", False, False),
        ("faithful leaders conduct themselves in the background precisely as they would if Jesus Christ or their senior pastor were seated in the front row", True, False),
        (". They do not manufacture synthetic devotion for public display while lapsing into negligence behind closed doors. Their diligence, pastoral care, and integrity remain consistent whether observed or unobserved, because their stewardship is rendered directly unto the Lord.", False, False)
    ])
    sp()

    # -------------------------------------------------------------------------
    # Section 2
    # -------------------------------------------------------------------------
    h("The Crippling Impact of Unfaithfulness and the High Value of Lives")
    ref("Proverbs 25:19")
    v("[19] Confidence in an unfaithful man in time of trouble is like a broken tooth, and a foot out of joint.")
    sp()

    b([
        ("The Apostle Paul's ministry was distinguished by faithful co-laborers, foremost of whom was Timothy. In epistles throughout the New Testament, Timothy appears consistently as Paul's trusted associate—a dependable son who naturally cared for the state of the flock. In stark contrast, ", False, False),
        ("Proverbs 25:19", True, False),
        (" articulates the acute damage caused by unreliable associates: ", False, False),
        ("*'Confidence in an unfaithful man in time of trouble is like a broken tooth, and a foot out of joint.'*", True, True),
        (" When one attempts to masticate with a fractured tooth, the result is excruciating pain and incapacity. When one attempts to walk on a dislocated joint, the entire body is crippled. In like manner, when an unfaithful individual occupies the office of a cell leader, department head, or pastor, their presence provides no functional stability; when ministerial pressure arrives, they collapse, inflicting agonizing hindrance upon the local church.", False, False)
    ])
    sp()

    b([
        ("Let every leader understand that serving God in the local church carries eternal weight and severe consequences. While believers rejoice in unconditional justification and the eternal grace of God, ministerial stewardship is not a trifling game. The Lord Jesus declared that it were better for a millstone to be hung about a man's neck and he be cast into the depth of the sea than to offend or cause one of His little ones to stumble (", False, False),
        ("Luke 17:1-2", True, False),
        ("). God places immense value upon human lives. When an individual handles God's flock flippantly, they sow seeds of ministerial dysfunction that will produce bitter fruit.", False, False)
    ])
    sp()

    b([
        ("*'I recall an individual who approached me concerning plans to relocate abroad. When I inquired what arrangements were being made for the disciples under his care, he responded with complete flippancy: \"Well, I was just thinking that since my pastor knows them, the church will take care of them.\" Such a casual, dismissive disposition toward God's sheep is spiritually disastrous. Individuals discard their pastoral responsibilities in Nigeria, arrive in North America or Europe, and then wonder why they encounter persistent frustration and sterility in ministry. They treated God's people like disposable commodities, dumping them into services without ongoing pastoral oversight until personal convenience dictated their departure.'*", False, True)
    ])
    sp()

    # -------------------------------------------------------------------------
    # Section 3
    # -------------------------------------------------------------------------
    h("The Currency of Eternity: Lives as Divine Capital and Investment")
    ref("Luke 16:10-12")
    v("[10] He that is faithful in that which is least is faithful also in much: and he that is unjust in the least is unjust also in much.")
    v("[11] If therefore ye have not been faithful in the unrighteous mammon, who will commit to your trust the true riches?")
    v("[12] And if ye have not been faithful in that which is another man's, who shall give you that which is your own?")
    sp()

    b([
        ("Human lives in the kingdom of God are analogous to capital investment. When God places ten, twenty, thirty, or fifty people under a cell leader's charge, He has entrusted that leader with spiritual currency of infinite value. Yet many leaders casually excuse numerical depletion, citing demanding employment, transit difficulties, or personal exhaustion, while simultaneously offering empty prayers for church increase.", False, False)
    ])
    sp()

    b([
        ("*'Consider a practical parallel: If you entrust an entrepreneur with one hundred thousand naira to trade and produce bean cakes (akara), and after weeks of silence, he returns claiming that because of societal disruptions or inflation, the entire principal is lost, but perhaps over the next year he might recover twenty thousand naira—would any prudent person commit a million naira to his hands? Any investor who does so would be acting with sheer foolishness. Yet leaders present an identical track record before God: ten souls were committed to their cell, nine have vanished through negligence with only a personal relative remaining, and yet they come to church conventions praying for increase. What increase can be committed to hands that squander divine investment?'*", False, True)
    ])
    sp()

    b([
        ("Every soul under a leader's care represents the accumulated spiritual investment of those who went before. The evangelist who preached, the pastor who taught, and the intercessors who labored all invested into that believer. For every soul a leader edifies and preserves, an entire generational lineage of disciples is impacted. Conversely, for every soul a leader neglects and loses, a branch of the kingdom's harvest is severed. When cell leaders fail to submit reports and offer trite excuses about depleted phone data or power outages, they demonstrate disregard for divine capital.", False, False)
    ])
    sp()

    b([
        ("Expose the hypocrisy of natural motivation: if a church administration were to announce a financial compensation of one thousand dollars for every soul won and successfully integrated into the local church, fatigue would evaporate overnight. Cell leaders would scour the streets, eliminate unnecessary sleep, overcome physical discomfort, and hold daily evangelistic outreaches with boundless energy. Yet the eternal value of a single human life bought by the blood of Christ infinitely surpasses earthly wealth. When leaders understand the true worth of souls, they labor past bodily fatigue, looking away from distracting symptoms and standing resolutely upon the authority of God's Word.", False, False)
    ])
    sp()

    # -------------------------------------------------------------------------
    # Section 4
    # -------------------------------------------------------------------------
    h("Success in Ministry Is Planned, Not Accidental")
    ref("1 Corinthians 9:26")
    v("[26] I therefore so run, not as uncertainly; so fight I, not as one that beateth the air.")
    sp()

    b([
        ("Ministers frequently make the erroneous assumption that numerical growth in other assemblies is merely the byproduct of doctrinal error, sensationalism, or theatrical gimmicks. While false teachings and charismatic manipulation abound in the religious landscape, numerical expansion across many works is actually built upon extraordinary dedication, rigorous work ethic, and meticulous administrative organization. Erroneous doctrine in itself does not gather people; discipline and labor do.", False, False)
    ])
    sp()

    b([
        ("*'Success in any endeavor is deliberate and planned. My father used to share that when he recognized his natural reading attention span did not exceed one hour, he did not surrender to mediocrity. Instead, he structured his day into five distinct one-hour study blocks separated by rest, eliminated extraneous social activities, and graduated with first-class honors. Spiritual ministry operates by the same principle. A cell group or local assembly will never flourish merely because the leader offers fervent prayers while neglecting practical execution. Praying without disciplined pastoral planning is like stockpiling ammunition in an armory without possessing a firearm to discharge it.'*", False, True)
    ])
    sp()

    b([
        ("Accountability must begin at the smallest scale. If a leadership core cannot meticulously track, follow up, and account for fifty attendees at a local convention, it is completely incapacitated from overseeing ten thousand or one hundred thousand souls. Faithfulness is the foundational breath of ministry: it requires an unbroken presence at one's duty post.", False, False)
    ])
    sp()

    b([
        ("Furthermore, the pastoral calling recognizes no secular vacations. Holidays declared by civil government are not holidays from the work of the ministry. While an employee may pause work, the spiritual steward recognizes that human needs, demonic pressures, and pastoral oversights operate continuously. In the secular world, emergency physicians, law enforcement personnel, and air traffic controllers do not abandon their posts on festive holidays; operations continue around the clock. If secular professionals exhibit such unyielding vigilance for temporal duties, how much more must the preacher of the eternal gospel maintain constant spiritual alertness?", False, False)
    ])
    sp()

    # -------------------------------------------------------------------------
    # Section 5
    # -------------------------------------------------------------------------
    h("The Doctrine of Divine Examination: Tested and Proven Character")
    ref("James 1:2-4")
    v("[2] My brethren, count it all joy when ye fall into divers temptations;")
    v("[3] Knowing this, that the trying of your faith worketh patience.")
    v("[4] But let patience have her perfect work, that ye may be perfect and entire, wanting nothing.")
    sp()

    b([
        ("God does not commit spiritual authority to untested vessels. When the prophet Samuel arrived to anoint a king in Bethlehem (", False, False),
        ("1 Samuel 16:6-7", True, False),
        ("), he looked upon Eliab's commanding physical stature and concluded that the Lord's anointed stood before him. But the Lord rejected Eliab, declaring that man looks on the outward appearance, but God looks upon the heart. In the original Hebrew, the term for rejection carries the force of one who has failed the moral examination. Divine testing is not designed to orchestrate human destruction; rather, it is akin to test-driving a newly engineered machine to demonstrate its structural integrity. God boasted of Job's resilience: ", False, False),
        ("*'Hast thou considered my servant Job, that there is none like him in the earth?'*", True, True),
        (" When trials confront the minister, James instructs that it should be received with joy, knowing that the trial of faith produces steadfast patience (", False, False),
        ("hypomone", True, True),
        ("), maturing the believer until they lack nothing.", False, False)
    ])
    sp()

    ref("1 Timothy 3:10, 13")
    v("[10] And let these also first be proved; then let them use the office of a deacon, being found blameless.")
    v("[13] For they that have used the office of a deacon well purchase to themselves a good degree, and great boldness in the faith which is in Christ Jesus.")
    sp()

    b([
        ("In ", False, False),
        ("1 Timothy 3:10", True, False),
        (", the Apostle Paul establishes a mandatory ecclesiastical statute: ", False, False),
        ("let these also first be proved", True, False),
        (". The term translated 'proved' is the Greek verb ", False, False),
        ("dokimazo", True, True),
        (", meaning to examine closely, scrutinize, and test thoroughly as metal in a furnace. Notice that Paul does not instruct Timothy to wait for a mystical sign from heaven; rather, human leadership is commanded to actively test and observe prospective workers. A leader must possess discerning eyes to assess individuals' character under pressure, their punctuality, their submission to instruction, and their dependability.", False, False)
    ])
    sp()

    ref("2 Timothy 2:15")
    v("[15] Study to shew thyself approved unto God, a workman that needeth not to be ashamed, rightly dividing the word of truth.")
    sp()

    b([
        ("Paul reinforces this standard in ", False, False),
        ("2 Timothy 2:15", True, False),
        (": the minister must labor to present themselves ", False, False),
        ("approved unto God", True, False),
        (". The word 'approved' is ", False, False),
        ("dokimos", True, True),
        ("—a certified, thoroughly tested workman who stands unashamed. Untested leadership is an ecclesiastical catastrophe. When the apostles in Jerusalem dispatched representatives to the Gentile churches in Antioch (", False, False),
        ("Acts 15:25-26", True, False),
        ("), their commendation was unequivocal: ", False, False),
        ("*'Men that have hazarded their lives for the name of our Lord Jesus Christ.'*", True, True),
        (" These were not fair-weather enthusiasts; their commitment had been demonstrated through persecution, sacrifice, and adversity.", False, False)
    ])
    sp()

    b([
        ("In the local church, genuine leadership promotion begins at the lowest, most unglamorous rungs of service. Individuals who join an assembly seeking rapid prominence, platforms, or ordination are promptly unmasked when assigned to church operations, sanitation, parking, or security. Those motivated by personal ambition flee when handed a mop or instructed to assist in the works department; true servants, however, count it an honor to serve anywhere they are placed, proving their hearts before God.", False, False)
    ])
    sp()

    # -------------------------------------------------------------------------
    # Section 6
    # -------------------------------------------------------------------------
    h("Joseph vs. Gehazi: The Archetypes of Ministerial Integrity")
    ref("Genesis 39:9")
    v("[9] How then can I do this great wickedness, and sin against God?")
    sp()

    b([
        ("When testing arrives, leaders invariably emerge as either a ", False, False),
        ("Joseph or a Gehazi", True, False),
        (". Joseph exemplified unwavering faithfulness across every station of life. In Canaan, he was the favored son; yet when cast into slavery in Potiphar's house, his diligence was so conspicuous that Potiphar committed the entire estate into his hands. When Potiphar's wife repeatedly solicited him to commit adultery, Joseph refused to compromise, asserting: ", False, False),
        ("*'How then can I do this great wickedness, and sin against God?'*", True, True),
        (" He fled the trap, leaving his garment behind.", False, False)
    ])
    sp()

    b([
        ("Falsely accused and thrown into Pharaoh's dungeon, Joseph did not collapse into self-pity, bitterness, or rebellion against God. Instead, he maintained his spiritual sensitivity, noticing the gloom on the faces of the baker and the butler and interpreting their prophetic dreams. When forgotten for two full years, he harbored no resentment. When summoned before Pharaoh, he did not negotiate terms, demand compensation for false imprisonment, or seek the execution of Potiphar's wife. He stood before the monarch with pure theological clarity, declaring that God would give Pharaoh an answer of peace.", False, False)
    ])
    sp()

    ref("Genesis 50:20")
    v("[20] But as for you, ye thought evil against me; but God meant it unto good, to bring to pass, as it is this day, to save much people alive.")
    sp()

    b([
        ("When exalted to govern Egypt, Joseph administered food to the nation, extending provision even to Potiphar's household without vindictiveness. When his brothers arrived trembling after their father's death, manufacturing an apologetic message out of fear, Joseph wept and comforted them, revealing the character of Christ: ", False, False),
        ("*'Fear not: for am I in the place of God? But as for you, ye thought evil against me; but God meant it unto good.'*", True, True),
        (" True faithfulness remains constant whether in the pit, the prison, or the palace.", False, False)
    ])
    sp()

    ref("2 Kings 5:26-27")
    v("[26] Is it a time to receive money, and to receive garments, and oliveyards, and vineyards, and sheep, and oxen, and menservants, and maidservants?")
    v("[27] The leprosy therefore of Naaman shall cleave unto thee, and unto thy seed for ever. And he went out from his presence a leper as white as snow.")
    sp()

    b([
        ("In tragic contrast stands Gehazi. Gehazi served under the prophet Elisha, witnessing supernatural demonstrations and holding the potential to inherit the prophetic succession. Yet when Naaman was healed of leprosy and Elisha refused material remuneration, Gehazi's hidden covetousness flared. Pursuing Naaman deceitfully, he fabricated a story to extract silver and garments. Elisha exposed his compromise: ", False, False),
        ("*'Went not mine heart with thee, when the man turned again from his chariot to meet thee?'*", True, True),
        (" Gehazi traded eternal spiritual inheritance for temporal gain, leaving Elisha's presence a leper as white as snow. Untested ambition invariably produces Gehazi-like catastrophes.", False, False)
    ])
    sp()

    b([
        ("Pastors and senior leaders must strictly observe this rule: ", False, False),
        ("never donate spiritual responsibility to unproven individuals", True, False),
        (". When individuals volunteer to launch independent groups or seek platforms without operating within established church structures, their ambition must be tested. God intentionally permits fires and difficult seasons so that those who cannot endure the desert are prevented from contaminating the promised land. Leadership oversight cannot be exercised through social media analytics; it requires direct, personal proving through the heat of service.", False, False)
    ])
    sp()

    # -------------------------------------------------------------------------
    # Section 7
    # -------------------------------------------------------------------------
    h("The Shepherd's Heart: David and the Defense of the Flock")
    ref("1 Samuel 16:11")
    v("[11] And Samuel said unto Jesse, Are here all thy children? And he said, There remaineth yet the youngest, and, behold, he keepeth the sheep.")
    sp()

    b([
        ("The scriptural portrait of shepherd leadership finds its pinnacle in David, whose life prophetically foreshadows Jesus Christ—the Good Shepherd (", False, False),
        ("Ezekiel 34:23; John 10:11", True, False),
        ("). When the prophet Samuel came to Bethlehem to anoint Israel's next king, all of Jesse's older, distinguished sons were presented, while David was absent in the wilderness. The scripture records the first mention of his character: ", False, False),
        ("*'Behold, he keepeth the sheep.'*", True, True),
        (" Faithfulness in obscurity is the prerequisite for public stewardship.", False, False)
    ])
    sp()

    ref("1 Samuel 17:28, 34-36")
    v("[28] And Eliab his eldest brother heard when he spake unto the men; and Eliab's anger was kindled against David, and he said... with whom hast thou left those few sheep in the wilderness?")
    v("[34] And David said unto Saul, Thy servant kept his father's sheep, and there came a lion, and a bear, and took a lamb out of the flock:")
    v("[35] And I went out after him, and smote him, and delivered it out of his mouth: and when he arose against me, I caught him by his beard, and smote him, and slew him.")
    v("[36] Thy servant slew both the lion and the bear: and this uncircumcised Philistine shall be as one of them.")
    sp()

    b([
        ("When David arrived at the battle encampment, his eldest brother Eliab falsely accused him of insolence, asking condescendingly: ", False, False),
        ("*'With whom hast thou left those few sheep in the wilderness?'*", True, True),
        (" But David had already ensured that those few sheep were placed in the custody of a dependable keeper (", False, False),
        ("1 Samuel 17:20", True, False),
        ("). Even when visiting a national battlefield, the safety of his small flock remained his primary concern.", False, False)
    ])
    sp()

    b([
        ("When brought before King Saul, who dismissed him as an untrained youth, David did not boast of political theories or military tactics. His credentials were born of pastoral vigilance: when a lion and a bear attacked the flock, David risked his own life, pursued the predators, struck them, delivered the lamb from their mouths, and slew them. He refused to allow a single lamb to perish. Moses tended the flock of Jethro for forty years in the back of the desert before shepherding Israel out of Egypt; David defended the few sheep before leading the twelve tribes. Faithfulness with the few is God's eternal measuring rod for greater kingdom stewardship.", False, False)
    ])
    sp()

    # -------------------------------------------------------------------------
    # Section 8
    # -------------------------------------------------------------------------
    h("The Indispensable Resources: The Word, the Spirit, and Spiritual Hunger")
    ref("1 Timothy 4:12")
    v("[12] Let no man despise thy youth; but be thou an example of the believers, in word, in conversation, in charity, in spirit, in faith, in purity.")
    sp()

    b([
        ("The cell leader's primary resources for multiplying divine capital are twofold: ", False, False),
        ("the Word of God and the Holy Spirit", True, False),
        (". The minister is not called to be a mere dispenser of arbitrary commands, but an ensample (", False, False),
        ("typos", True, True),
        (") to the flock in doctrine, speech, love, spiritual fervor, and moral purity (", False, False),
        ("1 Peter 5:3", True, False),
        ("). To lead effectively, a minister must cultivate ", False, False),
        ("the second hearing", True, False),
        (". A leader does not listen to biblical teaching merely for personal devotional encouragement; a leader listens on behalf of nations, absorbing doctrine, cross-references, and explanations to feed and establish those committed to their charge.", False, False)
    ])
    sp()

    b([
        ("*'A genuine leader possesses an insatiable, self-motivated appetite for God's Word. Decades ago, when biblical teaching literature and recordings were scarce in Nigeria, no one compelled me to acquire materials. I would trek through industrial quarters to Christian book distributors, inquire when new teaching volumes by Kenneth Hagin or Frederick K.C. Price were arriving, and reserve copies weeks in advance. When the Amplified Bible was extraordinarily rare, I would visit someone who owned a copy and spend thirty constrained minutes transcribing scripture text and cross-references by hand into notebooks. When teaching audio cassettes became available, I saved my financial resources and purchased entire library collections of over two hundred tapes, studying them meticulously. This was not platform posturing; it was an unquenchable hunger for the Word of God.'*", False, True)
    ])
    sp()

    b([
        ("*'We instituted correspondence study courses, distributing comprehensive study manuals to believers—including leaders who now pastor major branches—requiring them to answer detailed doctrinal questions so that truth was ingrained in their spirits. When denominations recently submitted an eighteen-hundred-page theological curriculum for review, expecting a four-month evaluation, I returned the completed critique within four days through relentless labor. In 1994, the Lord spoke clearly to my heart: \"Do not merely admire men's results; follow their labor.\" Pulpit effectiveness is not manufactured by theatrics or waving hands; it is the direct outcome of unseen hours spent in fasting, prayer, rigorous textual examination, and absolute submission to the Holy Spirit.'*", False, True)
    ])
    sp()

    b([
        ("This consecration extends to physical endurance and unwavering consistency. In the early days of ministry, believers walked long distances under intense sunlight to attend seminars, rejecting physical excuses and walking in divine healing. Leaders must rekindle that same apostolic stamina. Furthermore, spiritual consistency must become second nature—exemplified by senior ministers whose prayer routines in the Spirit are so clockwork and dependable that one can set their daily schedule by their fidelity. The Word and the Holy Spirit are the only resources that can sustain and multiply God's flock.", False, False)
    ])
    sp()

    # -------------------------------------------------------------------------
    # Section 9
    # -------------------------------------------------------------------------
    h("The Four Cardinal Demands of Pastoral Oversight")
    ref("Romans 15:20")
    v("[20] Yea, so have I strived to preach the gospel, not where Christ was named, lest I should build upon another man's foundation.")
    sp()

    b([
        ("In concluding the charge to cell leaders, the Apostle Paul's ministerial practice outlines ", False, False),
        ("four non-negotiable operational responsibilities", True, False),
        (" required of every cell leader, fellowship coordinator, and pastor:", False, False)
    ])
    sp()

    # Demand 1
    b([
        ("1. Pioneer Soul Winning (Evangelism): ", True, False),
        ("A cell leader must be an active, unyielding soul winner. Paul declared his apostolic rule: striving to preach the gospel where Christ was not previously named, refusing to build upon another man's foundation (", False, False),
        ("Romans 15:20", True, False),
        ("). The practice of poaching or re-circulating members from other churches is an ecclesiastical abomination that must never exist among us. True discipleship raises individuals from the foundation—winning the unsaved, reaching the unchurched, and establishing the untaught. In ", False, False),
        ("Numbers 27:17", True, False),
        (", Moses prayed for shepherds who would lead the congregation in and out. In ", False, False),
        ("Ezekiel 34:11", True, False),
        (", the true shepherd searches out the lost sheep. When seats in a cell meeting or church service remain empty, it is an unmistakable symptom of dormant soul winning. A cell leader must keep their soul-winning fire burning constantly.", False, False)
    ])
    sp()

    # Demand 2
    b([
        ("2. Active Establishment in the Local Church: ", True, False),
        ("It is utter ministerial negligence for a leader to dispatch people to church as detached spectators. A leader does not merely invite or send disciples; ", False, False),
        ("the leader attends church together with their disciples", True, False),
        (". Church attendance is an integral continuation of discipleship. It is completely irresponsible for a leader to sit comfortably in the front rows while their disciples are scattered aimlessly across the auditorium. Leaders must sit adjacent to the people they mentor, actively observing how they follow the Word, whether they open their Bibles, take notes, sing during worship, and pray with fervor. This is the essence of pastoral oversight (", False, False),
        ("episkope", True, True),
        ("). Just as negligent parenting produces undisciplined children who disrupt society, negligent pastoring produces careless believers. Leaders must provide direct, personal supervision.", False, False)
    ])
    sp()

    # Demand 3
    ref("Acts 19:9-10; Acts 20:31")
    v("[9] But when divers were hardened, and believed not... he departed from them, and separated the disciples, disputing daily in the school of one Tyrannus.")
    v("[10] And this continued by the space of two years; so that all they which dwelt in Asia heard the word of the Lord Jesus, both Jews and Greeks.")
    v("[31] Therefore watch, and remember, that by the space of three years I ceased not to warn every one night and day with tears.")
    sp()

    b([
        ("3. Daily Discipleship and Systematic Word Feeding: ", True, False),
        ("The work of Christian ministry is not a weekend recreational activity; it is a daily, unceasing vocation. In Ephesus, Paul gathered the disciples and taught them ", False, False),
        ("daily", True, False),
        (" for the space of two and a half years in the school of Tyrannus (", False, False),
        ("Acts 19:9-10", True, False),
        ("), testifying in ", False, False),
        ("Acts 20:31", True, False),
        (" that for three years he warned every one night and day with tears. While a cell leader cannot physically meet with every disciple daily, they must implement a systematic plan ensuring that every individual under their care feeds upon the Word of God daily. Cell leaders must track the sermon series and audio teachings their flock listens to, following up consistently: inquiring what was learned, discussing doctrinal principles, and ensuring continuous spiritual intake.", False, False)
    ])
    sp()

    # Demand 4
    ref("Acts 20:34-35")
    v("[34] Yea, ye yourselves know, that these hands have ministered unto my necessities, and to them that were with me.")
    v("[35] I have shewed you all things, how that so labouring ye ought to support the weak, and to remember the words of the Lord Jesus, how he said, It is more blessed to give than to receive.")
    sp()

    b([
        ("4. Ministering to Practical Needs and Maintaining Diligent Pastoral Records: ", True, False),
        ("Pastoral oversight encompasses both spiritual edification and practical care. Paul reminded the Ephesian elders that with his own hands he labored to minister to the necessities of those with him, supporting the weak (", False, False),
        ("Acts 20:34-35", True, False),
        ("). A shepherd must know the flock intimately. Every cell leader must maintain detailed, accurate pastoral records of everyone in their unit: their full legal names, dates of birth, residential addresses, occupations, and domestic circumstances. If a leader lacks the administrative discipline to maintain accurate biographical knowledge for twenty or fifty people, they fundamentally disqualify themselves from shepherding hundreds or thousands. Pastoral oversight requires deliberate, structured familiarity with every sheep.", False, False)
    ])
    sp()

    # -------------------------------------------------------------------------
    # Section 10
    # -------------------------------------------------------------------------
    h("Men of Purpose Are Busy People")
    b([
        ("Pastoral ministry is a continuous, daily lifestyle. Even when a leader is not physically present, they are pastoring through prayer, communication, strategic planning, and spiritual vigilance. Operating as an effective leader requires intentional time management: carving out disciplined boundaries for personal communion with God, family responsibilities, vocational employment, physical health, and flock oversight by ruthlessly eliminating frivolous, unproductive activities.", False, False)
    ])
    sp()

    b([
        ("*'Men of purpose are inherently busy people. When you observe leaders who accomplish monumental results in industry or ministry, their lives are characterized by high-velocity focus, decisive execution, and zero tolerance for idleness. One does not attain fulfillment in ministry through passive positive confessions while lounging in slothfulness. Fulfilling ministry demands diligent, relentless labor. Let your life be characterized by purpose. Stand as a dependable steward, wholly devoted to multiplying the investment of lives committed to your hands, until the Chief Shepherd appears.'*", False, True)
    ])
    sp()

    # Save Word docx
    for path in output_docx_paths:
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        doc.save(path)
        print(f"[+] Successfully saved Word docx to: {path}")

    # Save Markdown
    full_md_text = "\n".join(md_lines)
    for path in output_md_paths:
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(full_md_text)
        print(f"[+] Successfully saved Markdown to: {path}")

def count_words(doc_path):
    doc = Document(doc_path)
    total_words = 0
    for p in doc.paragraphs:
        words = p.text.split()
        total_words += len(words)
    return total_words

if __name__ == '__main__':
    base_dir = Path(r"c:\Users\Stephycopy\OneDrive\Desktop\Theos\theos")
    study_dir = base_dir / "Study group"

    docx_destinations = [
        base_dir / "Cell_Leaders_Conference_Faithfulness_in_Stewardship.docx",
        study_dir / "Cell_Leaders_Conference_Faithfulness_in_Stewardship.docx"
    ]

    md_destinations = [
        base_dir / "Cell_Leaders_Conference_Faithfulness_in_Stewardship.md",
        study_dir / "Cell_Leaders_Conference_Faithfulness_in_Stewardship.md"
    ]

    create_documents(docx_destinations, md_destinations)

    words = count_words(docx_destinations[0])
    print(f"[+] Verification: Total word count in Word document = {words} words.")
