#!/usr/bin/env python3
"""
format_faithfulness_in_ministry.py
Transforms the raw Fireflies transcript into a publication-grade document
following the Two Kinds of Righteousness Transcript Formatting Guide
(The Refined Claude Standard).
"""

import os
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def build_faithfulness_document(docx_paths, md_paths):
    doc = Document()

    # Page setup: Standard Letter (8.5" x 11.0"), 1.0" margins
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0x11, 0x18, 0x27) # Dark neutral

    md_lines = []

    def title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(18)
        md_lines.append(f"# {text}\n")

    def subtitle(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        md_lines.append(f"## *{text}*\n\n---\n")

    def h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"\n## {text}\n")
        return p

    def script_ref(ref_str):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(ref_str)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"> **{ref_str}**")
        return p

    def script_verse(verse_str):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(verse_str)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"> {verse_str}")
        return p

    def sp():
        # Space between blockquote and following text
        md_lines.append("")

    def p_body(runs):
        """
        runs is a list of tuples: (text, is_bold, is_italic)
        """
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        md_str = ""
        for text, bold, italic in runs:
            run = p.add_run(text)
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            if bold:
                run.bold = True
            if italic:
                run.italic = True

            # Markdown conversion
            if bold and italic:
                md_str += f"***{text}***"
            elif bold:
                md_str += f"**{text}**"
            elif italic:
                md_str += f"*{text}*"
            else:
                md_str += text
        md_lines.append(md_str + "\n")
        return p

    def p_anecdote(text):
        """Pure italic anecdote paragraph."""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        run = p.add_run(text)
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        run.italic = True
        md_lines.append(f"*{text}*\n")
        return p

    # =========================================================================
    # DOCUMENT ASSEMBLY (Refined Claude Standard)
    # =========================================================================

    title("Cell Leaders Conference — Faithfulness in Ministry")
    subtitle("December 2021 — Expository Teaching on Stewardship, Proven Character, and Pastoral Oversight")

    # -------------------------------------------------------------------------
    # SECTION 1
    # -------------------------------------------------------------------------
    h2("The Requirement of a Steward: Faithfulness, Reliability, and 100% Commitment to the Cause")

    script_ref("1 Corinthians 4:1-2")
    script_verse("[1] Let a man so account of us, as of the ministers of Christ, and stewards of the mysteries of God.")
    script_verse("[2] Moreover it is required in stewards, that a man be found faithful.")
    sp()

    p_body([
        ("In ", False, False),
        ("1 Corinthians 4:1-2", True, False),
        (", the Apostle Paul defines the governing baseline of Christian ministerial service. A Christian worker is not to be accounted by worldly titles, public acclaim, or social visibility, but as a minister of Christ and a steward of the mysteries of God. Paul declares the singular, non-negotiable requirement laid upon every steward: ", False, False),
        ("moreover it is required in stewards, that a man be found faithful", True, False),
        (". The word translated 'faithful' is the very same Greek word for faith, ", False, False),
        ("pistos", True, True),
        (", derived from the verb ", False, False),
        ("pisteuo", True, True),
        (". It denotes utter reliability, dependability, and steadfast trustworthiness under every condition.", False, False)
    ])

    script_ref("1 Corinthians 10:13")
    script_verse("[13] There hath no temptation taken you but such as is common to man: but God is faithful, who will not suffer you to be tempted above that ye are able...")
    sp()

    p_body([
        ("Believers frequently declare, 'God is faithful,' as affirmed in ", False, False),
        ("1 Corinthians 10:13", True, False),
        (". Why is God faithful? Because He is reliable, He is always available, and He never alters His word or reneges on His commitment in spite and despite opposing circumstances or extremities. Look at the intense pressure, hostility, and negativity God walked through leading up to the cross of Christ—yet He remained steadfast and never altered His word. That is why faithfulness is also called ", False, False),
        ("loyalty", True, False),
        (", and loyalty is defined as ", False, False),
        ("being committed to a cause 100%", True, True),
        (". When Paul asserts that it is required of a steward to be found faithful, he demands that every leader be found totally loyal to the assignment, the local church, and the spiritual oversight committed to their care.", False, False)
    ])

    # -------------------------------------------------------------------------
    # SECTION 2
    # -------------------------------------------------------------------------
    h2("Faithfulness vs. Ambition: Ruth and Naomi and the Diagnostic of Unfulfilled Desires")

    script_ref("Ruth 1:16-17")
    script_verse("[16] And Ruth said, Intreat me not to leave thee, or to return from following after thee: for whither thou goest, I will go; and where thou lodgest, I will lodge: thy people shall be my people, and thy God my God:")
    script_verse("[17] Where thou diest, will I die, and there will I be buried: the LORD do so to me, and more also, if ought but death part thee and me.")
    sp()

    p_body([
        ("Consider the covenant loyalty between Ruth and Naomi in ", False, False),
        ("Ruth 1", True, False),
        (". Following the death of her husband and sons in Moab, Naomi urged her daughters-in-law to return to their maternal homes. Orpah kissed Naomi and departed, but Ruth clave unto her, uttering that monumental confession: ", False, False),
        ("'Whither thou goest, I will go; and where thou lodgest, I will lodge: thy people shall be my people, and thy God my God.'", True, True),
        (" Ruth was an alien Gentile entering Israel with zero guarantees of financial security, marriage, or social status. She was looking for absolutely nothing in return. That is the very definition of faithfulness.", False, False)
    ])

    p_body([
        ("There is a sharp, irreconcilable difference between ", False, False),
        ("faithfulness and personal ambition", True, False),
        (". The definitive spiritual diagnostic that distinguishes the two is this: ", False, False),
        ("one discovers whether an individual is genuinely faithful or merely ambitious when their personal desires are delayed or denied", True, True),
        (". There are individuals who join a local church or enter ministry specifically to climb ecclesiastical ladders, gain prominence, or be noticed by senior leadership. Such people put on an elaborate appearance of faithfulness, but the moment their personal expectations suffer the slightest alteration, correction, or delay, their posture changes immediately. That sudden shift unmasks their true motive: they were never faithful. True faithfulness has nothing to do with favorable circumstances; faithful men and women do not break under pressure, make no excuses, and remain one hundred percent committed to the cause.", False, False)
    ])

    # -------------------------------------------------------------------------
    # SECTION 3
    # -------------------------------------------------------------------------
    h2("Wholehearted Servanthood: Elisha Pouring Water on Elijah's Hands and Wearing the Father's Ministry")

    script_ref("2 Kings 3:11")
    script_verse("[11] But Jehoshaphat said, Is there not here a prophet of the LORD, that we may enquire of the LORD by him? And one of the king of Israel's servants answered and said, Here is Elisha the son of Shaphat, which poured water on the hands of Elijah.")
    sp()

    p_body([
        ("The scriptural pattern of faithful stewardship is powerfully illustrated in Elijah and Elisha. Elisha did not know all that lay ahead when he was summoned from the plow; he simply followed Elijah with an undivided heart. Historical indications suggest Elisha came from an affluent agricultural household and was likely older, yet he willingly subjected himself to humble, menial servanthood. When the allied kings sought a prophet of the Lord, the accrediting testimony given concerning Elisha was: ", False, False),
        ("'Here is Elisha the son of Shaphat, which poured water on the hands of Elijah.'", True, True),
        (" He was celebrated not by oratorical titles, but by dependable, servant-hearted loyalty.", False, False)
    ])

    script_ref("2 Kings 2:14")
    script_verse("[14] And he took the mantle of Elijah that fell from him, and smote the waters, and said, Where is the LORD God of Elijah? and when he also had smitten the waters, they parted hither and thither: and Elisha went over.")
    sp()

    p_body([
        ("As Elijah made his final journey through Gilgal, Bethel, Jericho, and toward Jordan (", False, False),
        ("2 Kings 2", True, False),
        ("), the cynical sons of the prophets lingered behind as passive observers, questioning Elisha: 'Why are you following this master? Do you not know he will be taken from you today?' Elisha refused to be distracted by their cynicism and stood faithfully with his spiritual father to the very end. When Elijah asked what he desired, Elisha requested a double portion of his spirit—the legal birthright of the firstborn inheritor to maintain the continuity of the prophetic office. When Elijah was caught up and the mantle fell, Elisha demonstrated that his service was entirely devoid of personal ambition. Striking the waters of the Jordan, he cried: ", False, False),
        ("'Where is the LORD God of Elijah?'", True, True),
        (" He did not proclaim, 'Where is the God of Elisha?' He made it abundantly clear to all that he was merely wearing Elijah's ministry under the same God.", False, False)
    ])

    p_body([
        ("In identical fashion, Joshua honored Moses. He faithfully executed Moses' instructions and never exploited Moses' passing or human shortcomings as pulpit sermon material to manufacture an independent reputation. Faithful people do not conduct themselves differently when leadership is absent. Indeed, this establishes the cardinal rule of ministerial integrity: ", False, False),
        ("faithful leaders conduct themselves in the background precisely as they would if Jesus Christ or their senior pastor were seated in the front row", True, True),
        (". Their diligence, reverence, and pastoral care remain constant whether observed or unobserved.", False, False)
    ])

    # -------------------------------------------------------------------------
    # SECTION 4
    # -------------------------------------------------------------------------
    h2("The Broken Tooth and Foot out of Joint: The Acute Damage of Unfaithfulness")

    p_body([
        ("Consider the Apostle Paul's ministry. Throughout the New Testament epistles, Timothy appears consistently as Paul's dependable associate—so ubiquitous that one could readily refer to Paul's epistles as Timothy's epistles. Timothy was thoroughly reliable, sharing Paul's heart and naturally caring for the welfare of the flock.", False, False)
    ])

    script_ref("Proverbs 25:19")
    script_verse("[19] Confidence in an unfaithful man in time of trouble is like a broken tooth, and a foot out of joint.")
    sp()

    p_body([
        ("In sharp contrast, ", False, False),
        ("Proverbs 25:19", True, False),
        (" reveals the acute dysfunction and excruciating pain inflicted by unreliable workers: ", False, False),
        ("'Confidence in an unfaithful man in time of trouble is like a broken tooth, and a foot out of joint.'", True, True),
        (" When an individual attempts to chew with a fractured tooth, the result is sharp agony and incapacity. When one attempts to walk on a dislocated foot, mobility collapses. In like manner, when an unfaithful individual occupies the office of a cell leader, department head, or pastor, their presence provides no functional support; when ministerial pressure arrives, they collapse, crippling the local assembly and inflicting agonizing hindrance upon the work of God.", False, False)
    ])

    # -------------------------------------------------------------------------
    # SECTION 5
    # -------------------------------------------------------------------------
    h2("The Eternal Weight of Ministerial Consequences: God's Value for Human Lives")

    script_ref("Luke 17:1-2")
    script_verse("[1] It is impossible but that offences will come: but woe unto him, through whom they come!")
    script_verse("[2] It were better for him that a millstone were hanged about his neck, and he cast into the sea, than that he should offend one of these little ones.")
    sp()

    p_body([
        ("Let every leader understand that ministerial service carries severe spiritual consequences. While believers rejoice in unconditional justification and the eternal grace of God, ministry is not a casual game. Anyone who asserts that ministerial negligence carries no consequences is sorely deceived.", False, False)
    ])

    p_anecdote(
        "There was an individual I warned last year: 'Read my lips: before you were born, I have been teaching unconditional forgiveness. Before you were ever in danger of hell, I have been teaching the grace of God. But hear me clearly—you will bear the consequences of your actions.' He had done nothing to me personally, but I gave him that sober warning because when serving God, one cannot joke. God places immense value upon human lives. That is why the Lord Jesus declared in Luke 17:1-2 that it were better for a heavy millstone to be fastened around a man's neck and he be cast into the depth of the sea than to cause one of His little ones to stumble."
    )

    p_anecdote(
        "A church member once approached me sharing plans to relocate abroad. When I inquired, 'Have you thought about the disciples committed to your care?' he replied with complete flippancy: 'Oh, I was just thinking that since my pastor knows them, the church will take care of them.' I told him: 'No! You ought to be deeply concerned about those lives!' People discard their pastoral responsibilities in Nigeria, arrive in North America or Europe, and then wonder why they encounter persistent frustration and sterility in ministry. They treated God's people like disposable commodities, dumping them into services without ongoing pastoral oversight until personal convenience dictated their departure. One cannot think outside the people God has entrusted to their hands."
    )

    # -------------------------------------------------------------------------
    # SECTION 6
    # -------------------------------------------------------------------------
    h2("Lives as Divine Capital: The Akara Trader Analogy and the Soul-Winning Test")

    script_ref("Luke 16:10-12")
    script_verse("[10] He that is faithful in that which is least is faithful also in much: and he that is unjust in the least is unjust also in much.")
    script_verse("[11] If therefore ye have not been faithful in the unrighteous mammon, who will commit to your trust the true riches?")
    script_verse("[12] And if ye have not been faithful in that which is another man's, who shall give you that which is your own?")
    sp()

    p_body([
        ("In the kingdom of God, human lives are spiritual currency. When God places ten, twenty, thirty, or fifty people under a cell leader's charge, He has entrusted that leader with capital of infinite eternal worth. Yet many leaders casually excuse numerical decline, citing secular employment, transit difficulties, or residential distance, while simultaneously offering empty prayers for church increase.", False, False)
    ])

    p_anecdote(
        "Consider the practical illustration I shared yesterday: If you entrust an entrepreneur with one hundred thousand naira to trade and fry akara (bean cakes), and after weeks of silence he returns claiming, 'COVID came, people are wearing masks and not eating akara, beans are now expensive, and the entire principal is lost, but to the glory of God, I might recover twenty thousand naira over the next year'—would any prudent person commit a million naira into his hands? Anyone who gives him fifty thousand naira is an idiot! Look at his track record. Yet leaders present an identical track record before God: ten souls were committed to their cell, nine have vanished through negligence with only a personal cousin remaining, and yet they come to church conventions praying for increase! More of what? What increase can be committed to hands that squander divine investment?"
    )

    p_body([
        ("Every soul under a leader's care represents the accumulated spiritual investment of those who went before—the evangelist who proclaimed the gospel, the pastor who taught, and the intercessors who labored. For every soul a leader edifies and preserves, a generational lineage of disciples is impacted. Conversely, for every soul a leader neglects and loses, a branch of the kingdom's harvest is severed. When cell leaders fail to submit daily reports and offer trite excuses about depleted data or power outages, they demonstrate disregard for divine capital.", False, False)
    ])

    p_body([
        ("Expose the hypocrisy of natural motivation: if church leadership were to announce a financial compensation of one thousand dollars for every soul won and successfully integrated into the local assembly, streets would be covered within a month! For ten souls, a leader would receive ten thousand dollars; for one hundred souls, one hundred thousand dollars. Complaints of evening fatigue would vanish instantly; leaders would claim divine strength, eliminate unnecessary sleep, and hold continuous outreaches because monetary gain is in view. Yet the eternal value of a single human life bought by the blood of Christ infinitely surpasses earthly wealth.", False, False)
    ])

    p_anecdote(
        "Recently, I experienced a severe, troubling headache. Someone in the room urged: 'Pastor, you need to rest.' I replied: 'No. I know I should rest, but there is work to be done. I will look away from physical symptoms and act like the Word of God is true. The job has to be done.' Souls are precious; their value demands unwavering consecration."
    )

    # -------------------------------------------------------------------------
    # SECTION 7
    # -------------------------------------------------------------------------
    h2("Success in Ministry Is Planned: Dedication, Structure, and Ministerial Alertness")

    script_ref("1 Corinthians 9:26")
    script_verse("[26] I therefore so run, not as uncertainly; so fight I, not as one that beateth the air.")
    sp()

    p_body([
        ("Ministers frequently make the erroneous assumption that numerical growth in other assemblies is merely the byproduct of false doctrine, sensationalism, or theatrical gimmicks like mantles and holy oil. That is completely false. In many instances, those assemblies expand because their leaders are totally dedicated, exceptionally hardworking, and administratively disciplined. Erroneous doctrine in itself does not gather people; discipline and labor do.", False, False)
    ])

    p_anecdote(
        "My father taught me that success in any endeavor is deliberate and planned. When he recognized that his natural reading attention span did not exceed one hour, he did not surrender to mediocrity. Instead, he structured his day into five distinct one-hour study blocks separated by rest, eliminated extraneous social activities, and graduated with first-class honors. Spiritual ministry operates by the very same principle. A cell group or local assembly will never flourish merely because a leader offers fervent prayers while neglecting practical execution. Praying without disciplined pastoral planning is like stockpiling ammunition in an armory without possessing a firearm to discharge it."
    )

    p_body([
        ("Accountability must begin at the smallest scale. If a leadership core cannot meticulously track, follow up, and account for attendees at a convention, it is completely incapacitated from overseeing tens of thousands or hundreds of thousands of souls. Faithfulness is the foundational breath of ministry: it requires an unbroken presence at one's duty post.", False, False)
    ])

    p_body([
        ("Furthermore, the pastoral calling recognizes no secular vacations. Holidays declared by civil government are not holidays from the work of the ministry. While temporal employees pause work, the spiritual steward recognizes that human needs, demonic pressures, and pastoral oversights operate continuously.", False, False)
    ])

    p_anecdote(
        "One Christmas Day, returning from the United States via Dallas, I had an extended conversation with Brother Pelumi and missed my flight due to casual inattention. Later, observing from the airplane, I noticed airport security working, planes flying, air traffic controllers directing flights without singing holiday carols, emergency physicians on call delivering babies on Christmas Day, and police officers inspecting vehicle trunks while wishing motorists a Merry Christmas. Operations were continuing around the clock. I told myself: 'I am a preacher of the eternal gospel; I must maintain unyielding spiritual alertness.' Faithful leaders must be faithful with time and alertness in the execution of their duties."
    )

    # -------------------------------------------------------------------------
    # SECTION 8
    # -------------------------------------------------------------------------
    h2("Proving and Testing: The Biblical Ecclesiastical Statute for Leadership")

    script_ref("1 Samuel 16:6-7")
    script_verse("[6] And it came to pass, when they were come, that he looked on Eliab, and said, Surely the LORD'S anointed is before him.")
    script_verse("[7] But the LORD said unto Samuel, Look not on his countenance, or on the height of his stature; because I have refused him: for the LORD seeth not as man seeth; for man looketh on the outward appearance, but the LORD looketh on the heart.")
    sp()

    p_body([
        ("God does not commit spiritual authority to untested vessels. When the prophet Samuel arrived to anoint a king in Bethlehem (", False, False),
        ("1 Samuel 16:6-7", True, False),
        ("), he looked upon Eliab's commanding physical stature and concluded that the Lord's anointed stood before him. But the Lord rejected Eliab, declaring that man looks on the outward appearance, but God looks upon the heart. In the original Hebrew, the term for rejection carries the force of one who has failed the moral examination. Divine testing is not designed to orchestrate human destruction; rather, it is akin to test-driving a newly engineered machine to demonstrate its structural integrity.", False, False)
    ])

    script_ref("James 1:2-4")
    script_verse("[2] My brethren, count it all joy when ye fall into divers temptations;")
    script_verse("[3] Knowing this, that the trying of your faith worketh patience.")
    script_verse("[4] But let patience have her perfect work, that ye may be perfect and entire, wanting nothing.")
    sp()

    p_body([
        ("God boasted of Job's integrity: ", False, False),
        ("'Hast thou considered my servant Job, that there is none like him in the earth?'", True, True),
        (" When trials confront the minister, James instructs that it should be received with joy, knowing that the trial of faith produces steadfast patience (", False, False),
        ("hypomone", True, True),
        ("), maturing the believer until they lack nothing. Running from divine tests is running from the very character required to fulfill the ministry.", False, False)
    ])

    script_ref("1 Timothy 3:10, 13")
    script_verse("[10] And let these also first be proved; then let them use the office of a deacon, being found blameless.")
    script_verse("[13] For they that have used the office of a deacon well purchase to themselves a good degree, and great boldness in the faith which is in Christ Jesus.")
    sp()

    p_body([
        ("In ", False, False),
        ("1 Timothy 3:10", True, False),
        (", the Apostle Paul establishes a mandatory ecclesiastical statute: ", False, False),
        ("let these also first be proved", True, False),
        (". The term translated 'proved' is the Greek verb ", False, False),
        ("dokimazo", True, True),
        (", meaning to scrutinize, examine closely, and test thoroughly as metal in a furnace. Notice that Paul does not instruct Timothy to wait for a mystical sign from heaven; rather, human leadership is commanded to actively test and observe prospective workers. A leader must possess discerning eyes to assess individuals' character under pressure, their punctuality, their submission to instruction, and their dependability.", False, False)
    ])

    script_ref("2 Timothy 2:15")
    script_verse("[15] Study to shew thyself approved unto God, a workman that needeth not to be ashamed, rightly dividing the word of truth.")
    sp()

    p_body([
        ("Paul reinforces this standard in ", False, False),
        ("2 Timothy 2:15", True, False),
        (": the minister must labor to present themselves ", False, False),
        ("approved unto God", True, False),
        (". The word 'approved' is ", False, False),
        ("dokimos", True, True),
        ("—a certified, thoroughly tested workman who stands unashamed. Untested leadership is an ecclesiastical catastrophe.", False, False)
    ])

    script_ref("Acts 15:25-26")
    script_verse("[25] It seemed good unto us, being assembled with one accord, to send chosen men unto you with our beloved Barnabas and Paul,")
    script_verse("[26] Men that have hazarded their lives for the name of our Lord Jesus Christ.")
    sp()

    p_body([
        ("When the apostles in Jerusalem dispatched representatives to the Gentile assemblies (", False, False),
        ("Acts 15:25-26", True, False),
        ("), their commendation was unequivocal: ", False, False),
        ("'Men that have hazarded their lives for the name of our Lord Jesus Christ.'", True, True),
        (" These were not fair-weather enthusiasts; their commitment had been demonstrated through persecution, sacrifice, and adversity.", False, False)
    ])

    p_anecdote(
        "Certain individuals join a local church seeking rapid prominence, platforms, or ordination. When individuals arrive boasting of previous titles—even if they were an archbishop elsewhere—they must pocket the title at the door, attend membership classes, and be tested in Operations, Works, and Security. Those who came desiring only to sing in the choir were assigned to clean and watch the toilets. True servants count it an honor to serve anywhere they are placed; those motivated by carnal ambition flee when handed a mop or instructed to assist in menial tasks."
    )

    # -------------------------------------------------------------------------
    # SECTION 9
    # -------------------------------------------------------------------------
    h2("Joseph vs. Gehazi: Integrity Under Pressure vs. Covetous Compromise")

    script_ref("Genesis 39:9")
    script_verse("[9] There is none greater in this house than I; neither hath he kept back any thing from me but thee, because thou art his wife: how then can I do this great wickedness, and sin against God?")
    sp()

    p_body([
        ("When testing arrives, leaders invariably emerge as either a ", False, False),
        ("Joseph or a Gehazi", True, False),
        (". Joseph exemplified unwavering faithfulness across every station of life. In Canaan, he was his father's favored son; yet when cast into slavery in Potiphar's house, his diligence was so conspicuous that Potiphar committed the entire estate into his hands. When Potiphar's wife repeatedly solicited him to commit adultery, Joseph refused to compromise, asserting: ", False, False),
        ("'How then can I do this great wickedness, and sin against God?'", True, True),
        (" He fled the trap, leaving his garment behind.", False, False)
    ])

    p_body([
        ("Falsely accused and thrown into Pharaoh's dungeon, Joseph did not collapse into self-pity, bitterness, or rebellion against God. Instead, he maintained his spiritual sensitivity, interpreting the prophetic dreams of the baker and the butler. When forgotten for two full years, he harbored no resentment. When summoned before Pharaoh, he did not negotiate terms, demand compensation for false imprisonment, or seek the execution of Potiphar's wife. He stood before the monarch with pure theological clarity, declaring that God would give Pharaoh an answer of peace.", False, False)
    ])

    script_ref("Genesis 50:20")
    script_verse("[20] But as for you, ye thought evil against me; but God meant it unto good, to bring to pass, as it is this day, to save much people alive.")
    sp()

    p_body([
        ("When exalted to govern Egypt, Joseph administered food to the nation, extending provision even to Potiphar's household without vindictiveness. When his brothers arrived trembling after Jacob's death, manufacturing an apologetic message out of fear, Joseph wept and comforted them, revealing the character of Christ: ", False, False),
        ("'Fear not: for am I in the place of God? But as for you, ye thought evil against me; but God meant it unto good.'", True, True),
        (" True faithfulness remains constant whether in the pit, the prison, or the palace.", False, False)
    ])

    script_ref("2 Kings 5:26-27")
    script_verse("[26] And he said unto him, Went not mine heart with thee, when the man turned again from his chariot to meet thee? Is it a time to receive money, and to receive garments, and oliveyards, and vineyards, and sheep, and oxen, and menservants, and maidservants?")
    script_verse("[27] The leprosy therefore of Naaman shall cleave unto thee, and unto thy seed for ever. And he went out from his presence a leper as white as snow.")
    sp()

    p_body([
        ("In tragic contrast stands Gehazi. Gehazi served under the prophet Elisha, witnessing supernatural demonstrations and holding the potential to inherit the prophetic succession. Yet when Naaman was healed of leprosy and Elisha refused material remuneration, Gehazi's hidden covetousness flared. Pursuing Naaman deceitfully, he fabricated a story to extract silver and garments. Elisha exposed his compromise: ", False, False),
        ("'Went not mine heart with thee, when the man turned again from his chariot to meet thee?'", True, True),
        (" Gehazi traded eternal spiritual inheritance for temporal gain, leaving Elisha's presence a leper as white as snow. Untested ambition invariably produces Gehazi-like catastrophes.", False, False)
    ])

    p_body([
        ("Pastors and senior leaders must strictly observe this rule: ", False, False),
        ("never donate spiritual responsibility to unproven individuals", True, False),
        (". When individuals volunteer to launch independent groups or seek platforms without operating within established church structures, their ambition must be tested. God intentionally permits fires and difficult seasons so that those who cannot endure the desert are prevented from contaminating the promised land. Leadership oversight cannot be exercised through social media analytics; it requires direct, personal proving through the heat of service.", False, False)
    ])

    # -------------------------------------------------------------------------
    # SECTION 10
    # -------------------------------------------------------------------------
    h2("Keeping the Few Sheep: David's Pastoral Faithfulness in Obscurity")

    script_ref("1 Samuel 16:11")
    script_verse("[11] And Samuel said unto Jesse, Are here all thy children? And he said, There remaineth yet the youngest, and, behold, he keepeth the sheep. And Samuel said unto Jesse, Send and fetch him: for we will not sit down till he come hither.")
    sp()

    p_body([
        ("The scriptural portrait of shepherd leadership finds its pinnacle in David, whose life prophetically foreshadows Jesus Christ—the Good Shepherd (", False, False),
        ("Ezekiel 34:23; John 10:11", True, False),
        ("). When the prophet Samuel arrived in Bethlehem to anoint Israel's next king, all of Jesse's older sons were presented, while David was absent in the wilderness. The scripture records the first mention of his character: ", False, False),
        ("'Behold, he keepeth the sheep.'", True, True),
        (" Faithfulness in obscurity is the prerequisite for public stewardship.", False, False)
    ])

    script_ref("1 Samuel 17:20, 28, 34-36")
    script_verse("[20] And David rose up early in the morning, and left the sheep with a keeper...")
    script_verse("[28] And Eliab his eldest brother heard when he spake unto the men; and Eliab's anger was kindled against David, and he said, Why camest thou down hither? and with whom hast thou left those few sheep in the wilderness?")
    script_verse("[34] And David said unto Saul, Thy servant kept his father's sheep, and there came a lion, and a bear, and took a lamb out of the flock:")
    script_verse("[35] And I went out after him, and smote him, and delivered it out of his mouth: and when he arose against me, I caught him by his beard, and smote him, and slew him.")
    script_verse("[36] Thy servant slew both the lion and the bear: and this uncircumcised Philistine shall be as one of them...")
    sp()

    p_body([
        ("When David arrived at the battle encampment, his eldest brother Eliab falsely accused him of insolence, asking condescendingly: ", False, False),
        ("'With whom hast thou left those few sheep in the wilderness?'", True, True),
        (" But David had already ensured that those few sheep were placed in the custody of a dependable keeper (", False, False),
        ("1 Samuel 17:20", True, False),
        ("). Even when visiting a national battlefield, the safety of his small flock remained his primary concern.", False, False)
    ])

    p_body([
        ("When brought before King Saul, who dismissed him as an untrained youth, David did not boast of political theories or military tactics. His credentials were born of pastoral vigilance: when a lion and a bear attacked the flock, David risked his own life, pursued the predators, struck them, delivered the lamb from their mouths, and slew them. He refused to allow a single lamb to perish. Moses tended the flock of Jethro for forty years in the back of the desert before shepherding Israel out of Egypt; David defended the few sheep before leading the twelve tribes. Faithfulness with the few is God's eternal measuring rod for greater kingdom stewardship.", False, False)
    ])

    # -------------------------------------------------------------------------
    # SECTION 11
    # -------------------------------------------------------------------------
    h2("The Word and the Spirit: Cultivating the Second Hearing and Unquenchable Hunger")

    script_ref("1 Timothy 4:12")
    script_verse("[12] Let no man despise thy youth; but be thou an example of the believers, in word, in conversation, in charity, in spirit, in faith, in purity.")
    script_ref("1 Peter 5:3")
    script_verse("[3] Neither as being lords over God's heritage, but being ensamples to the flock.")
    sp()

    p_body([
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

    p_anecdote(
        "A genuine leader possesses an insatiable, self-motivated appetite for God's Word. Decades ago, when biblical teaching literature and recordings were scarce in Nigeria, no one compelled me to acquire materials. I would trek through industrial quarters to Christian book distributors, inquire when new teaching volumes by Kenneth Hagin or Frederick K.C. Price were arriving, and reserve copies weeks in advance. When the Amplified Bible was extraordinarily rare, I would visit someone who owned a copy and spend thirty constrained minutes transcribing scripture text and cross-references by hand into notebooks. When teaching audio cassettes became available, I saved my financial resources and purchased entire library collections of over two hundred tapes, studying them meticulously. This was not platform posturing; it was an unquenchable hunger for the Word of God."
    )

    p_anecdote(
        "We instituted correspondence study courses, distributing comprehensive study manuals to believers—including leaders who now pastor major branches—requiring them to answer detailed doctrinal questions so that truth was ingrained in their spirits. When denominations recently submitted an eighteen-hundred-page theological curriculum for review, expecting a four-month evaluation, I returned the completed critique within four days through relentless labor. In 1994, the Lord spoke clearly to my heart: 'Do not merely admire men's results; follow their labor.' Pulpit effectiveness is not manufactured by theatrics or waving hands; it is the direct outcome of unseen hours spent in fasting, prayer, rigorous textual examination, and absolute submission to the Holy Spirit."
    )

    p_body([
        ("This consecration extends to physical endurance and unwavering consistency. In the early days of ministry, believers walked long distances under intense sunlight to attend seminars, rejecting physical excuses and walking in divine healing. Leaders must rekindle that same apostolic stamina. Furthermore, spiritual consistency must become second nature—exemplified by senior ministers whose prayer routines in the Spirit are so clockwork and dependable that one can set their daily schedule by their fidelity. The Word and the Holy Spirit are the only resources that can sustain and multiply God's flock.", False, False)
    ])

    # -------------------------------------------------------------------------
    # SECTION 12
    # -------------------------------------------------------------------------
    h2("The Four Demands of Pastoral Oversight: Evangelism, Establishment, Daily Feeding, and Pastoral Care")

    script_ref("Romans 15:20")
    script_verse("[20] Yea, so have I strived to preach the gospel, not where Christ was named, lest I should build upon another man's foundation.")
    sp()

    p_body([
        ("In concluding the charge to cell leaders, the Apostle Paul's ministerial practice outlines ", False, False),
        ("four non-negotiable operational responsibilities", True, False),
        (" required of every cell leader, fellowship coordinator, and pastor:", False, False)
    ])

    p_body([
        ("1. Pioneer Soul Winning (Evangelism): ", True, False),
        ("A cell leader must be an active, unyielding soul winner. Paul declared his apostolic rule: striving to preach the gospel where Christ was not previously named, refusing to build upon another man's foundation (", False, False),
        ("Romans 15:20", True, False),
        ("). The practice of poaching or re-circulating members from other churches is an ecclesiastical abomination that must never exist among us. True discipleship raises individuals from the foundation—winning the unsaved, reaching the unchurched, and establishing the untaught. In ", False, False),
        ("Numbers 27:17", True, False),
        (", Moses prayed for shepherds who would lead the congregation in and out. In ", False, False),
        ("Ezekiel 34:11", True, False),
        (", the true shepherd searches out the lost sheep. When seats in a cell meeting or church service remain empty, it is an unmistakable symptom of dormant soul winning. A cell leader must keep their soul-winning fire burning constantly.", False, False)
    ])

    p_body([
        ("2. Active Establishment in the Local Church: ", True, False),
        ("It is utter ministerial negligence for a leader to dispatch people to church as detached spectators. A leader does not merely invite or send disciples; ", False, False),
        ("the leader attends church together with their disciples", True, False),
        (". Church attendance is an integral continuation of discipleship. It is completely irresponsible for a leader to sit comfortably in the front rows while their disciples are scattered aimlessly across the auditorium. Leaders must sit adjacent to the people they mentor, actively observing how they follow the Word, whether they open their Bibles, take notes, sing during worship, and pray with fervor. This is the essence of pastoral oversight (", False, False),
        ("episkope", True, True),
        ("). Just as negligent parenting produces undisciplined children who disrupt society, negligent pastoring produces careless believers. Leaders must provide direct, personal supervision.", False, False)
    ])

    script_ref("Acts 19:9-10")
    script_verse("[9] But when divers were hardened, and believed not, but spake evil of that way before the multitude, he departed from them, and separated the disciples, disputing daily in the school of one Tyrannus.")
    script_verse("[10] And this continued by the space of two years; so that all they which dwelt in Asia heard the word of the Lord Jesus, both Jews and Greeks.")
    script_ref("Acts 20:31")
    script_verse("[31] Therefore watch, and remember, that by the space of three years I ceased not to warn every one night and day with tears.")
    sp()

    p_body([
        ("3. Daily Discipleship and Systematic Word Feeding: ", True, False),
        ("The work of Christian ministry is not a weekend recreational activity; it is a daily, unceasing vocation. In Ephesus, Paul gathered the disciples and taught them ", False, False),
        ("daily", True, False),
        (" for the space of two and a half years in the school of Tyrannus (", False, False),
        ("Acts 19:9-10", True, False),
        ("), testifying in ", False, False),
        ("Acts 20:31", True, False),
        (" that for three years he warned every one night and day with tears. While a cell leader cannot physically meet with every disciple daily, they must implement a systematic plan ensuring that every individual under their care feeds upon the Word of God daily. Cell leaders must track the sermon series and audio teachings their flock listens to, following up consistently: inquiring what was learned, discussing doctrinal principles, and ensuring continuous spiritual intake.", False, False)
    ])

    script_ref("Acts 20:34-35")
    script_verse("[34] Yea, ye yourselves know, that these hands have ministered unto my necessities, and to them that were with me.")
    script_verse("[35] I have shewed you all things, how that so labouring ye ought to support the weak, and to remember the words of the Lord Jesus, how he said, It is more blessed to give than to receive.")
    sp()

    p_body([
        ("4. Ministering to Practical Needs and Maintaining Diligent Pastoral Records: ", True, False),
        ("Pastoral oversight encompasses both spiritual edification and practical care. Paul reminded the Ephesian elders that with his own hands he labored to minister to the necessities of those with him, supporting the weak (", False, False),
        ("Acts 20:34-35", True, False),
        ("). A shepherd must know the flock intimately. Every cell leader must maintain detailed, accurate pastoral records of everyone in their unit: their full legal names, dates of birth, residential addresses, occupations, and domestic circumstances. If a leader lacks the administrative discipline to maintain accurate biographical knowledge for twenty or fifty people, they fundamentally disqualify themselves from shepherding hundreds or thousands. Pastoral oversight requires deliberate, structured familiarity with every sheep.", False, False)
    ])

    # -------------------------------------------------------------------------
    # SECTION 13
    # -------------------------------------------------------------------------
    h2("Men of Purpose Are Busy People: Relentless Labor and Faithful Fulfillment")

    p_body([
        ("Pastoral ministry is a continuous, daily lifestyle. Even when a leader is not physically present, they are pastoring through prayer, communication, strategic planning, and spiritual vigilance. Operating as an effective leader requires intentional time management: carving out disciplined boundaries for personal communion with God, family responsibilities, vocational employment, physical health, and flock oversight by ruthlessly eliminating frivolous, unproductive activities.", False, False)
    ])

    p_anecdote(
        "Men of purpose are inherently busy people. When you observe leaders who accomplish monumental results in industry or ministry, their lives are characterized by high-velocity focus, decisive execution, and zero tolerance for idleness. One does not attain fulfillment in ministry through passive positive confessions while lounging in slothfulness. Fulfilling ministry demands diligent, relentless labor. Let your life be characterized by purpose. Stand as a dependable steward, wholly devoted to multiplying the investment of lives committed to your hands, until the Chief Shepherd appears."
    )

    # -------------------------------------------------------------------------
    # EXPORT FILES
    # -------------------------------------------------------------------------
    # Save Word DOCX
    for docx_p in docx_paths:
        Path(docx_p).parent.mkdir(parents=True, exist_ok=True)
        doc.save(docx_p)
        print(f"[+] Successfully saved Word docx to: {docx_p}")

    # Save Markdown
    full_md = "\n".join(md_lines)
    for md_p in md_paths:
        Path(md_p).parent.mkdir(parents=True, exist_ok=True)
        with open(md_p, "w", encoding="utf-8") as f:
            f.write(full_md)
        print(f"[+] Successfully saved Markdown to: {md_p}")

if __name__ == '__main__':
    base_dir = Path(r"c:\Users\Stephycopy\OneDrive\Desktop\Theos\theos")
    study_dir = base_dir / "Study group"

    docx_files = [
        base_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry.docx",
        study_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry.docx"
    ]

    md_files = [
        base_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry.md",
        study_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry.md"
    ]

    build_faithfulness_document(docx_files, md_files)

    # Count words
    doc = Document(docx_files[0])
    total_words = sum(len(p.text.split()) for p in doc.paragraphs)
    print(f"[+] Document created. Total Word count = {total_words} words.")
