#!/usr/bin/env python3
"""
build_power_gifts_track2.py
Compiles the comprehensive, verbatim transcript of:
"In the Spirit by the Spirit – Power Gifts - Track 2"
Taught by Pastor Chris Segun Onayinka (Saints Community Church).

Adheres strictly to the Refined Claude Standard Verbatim Preservation:
- 100% preservation of the speaker's authentic vocabulary, definitions, and concepts.
- Complete removal of oral fillers ('praise the Lord', 'are you with me', 'look at me please', 'huh?').
- Smoothing of oral syntax into publication-grade expository paragraphs.
- Pronoun harmonization into consistent third-person exposition.
- Scripture citations formatted into bold blockquotes.
- Anecdotes, illustrations, and personal vignettes formatted in italics.
- Document output in DOCX (Arial 12pt, 1.15 line spacing, 1-inch margins) and Markdown (.md).
"""

from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def build_power_gifts_transcript(docx_paths, md_paths, raw_path=None):
    doc = Document()

    # Standard Letter (8.5" x 11.0"), 1.0" margins
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
    font.color.rgb = RGBColor(0x11, 0x18, 0x27)

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

    def h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"\n### {text}\n")
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
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(verse_str)
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"> {verse_str}")
        return p

    def p_body(runs):
        if md_lines and md_lines[-1].startswith(">"):
            md_lines.append("")
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(6)
        line_md = []
        for text, bold, italic in runs:
            r = p.add_run(text)
            r.bold = bold
            r.italic = italic
            r.font.name = 'Arial'
            r.font.size = Pt(12)
            if bold and italic:
                line_md.append(f"***{text}***")
            elif bold:
                line_md.append(f"**{text}**")
            elif italic:
                line_md.append(f"*{text}*")
            else:
                line_md.append(text)
        md_lines.append("".join(line_md) + "\n")
        return p

    def p_anecdote(text):
        if md_lines and md_lines[-1].startswith(">"):
            md_lines.append("")
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run(text)
        r.italic = True
        r.font.name = 'Arial'
        r.font.size = Pt(12)
        md_lines.append(f"*{text}*\n")
        return p

    # --- DOCUMENT HEADER ---
    title("In the Spirit, By the Spirit — Power Gifts (Track 2)")
    subtitle("Teaching by Pastor Chris Segun Onayinka — Saints Community Church")

    # --- SECTION 1 ---
    h2("1. Salvation and the Giving of the Spirit: The Indwelling Seal of God")

    p_body([
        ("In the teaching series examining the things of the Spirit, the starting point of understanding rests upon the foundational relationship between salvation and the Holy Spirit. Paul writes to the Corinthians regarding the faithfulness and established promises of God:", False, False)
    ])

    script_ref("2 Corinthians 1:19-22")
    script_verse("[19] For the Son of God, Jesus Christ, who was preached among you by us, even by me and Silvanus and Timotheus, was not yea and nay, but in him was yea.")
    script_verse("[20] For all the promises of God in him are yea, and in him Amen, unto the glory of God by us.")
    script_verse("[21] Now he which stablisheth us with you in Christ, and hath anointed us, is God;")
    script_verse("[22] Who hath also sealed us, and given the earnest of the Spirit in our hearts.")

    p_body([
        ("In Romans 15:8, the apostle similarly confirms that Christ confirmed the promises made unto the fathers:", False, False)
    ])

    script_ref("Romans 15:8")
    script_verse("[8] Now I say that Jesus Christ was a minister of the circumcision for the truth of God, to confirm the promises made unto the fathers.")

    p_body([
        ("Salvation is completely synonymous with the giving of the Spirit. There cannot be salvation without the giving of the Spirit; the giving of the Spirit is salvation. One cannot be saved and not have the Spirit of God. To receive salvation is to receive the Spirit of God as the gift of God. The terms are interchangeable in New Testament theology: the gift of the Spirit is salvation.", False, False)
    ])

    # --- SECTION 2 ---
    h2("2. The Three Events of the Gospel: Why Resurrection Is the Crown of Redemption")

    p_body([
        ("A common misunderstanding among believers is to restrict the sacrifice of Jesus exclusively to the cross, assuming his death is the entirety of salvation. In biblical theology, the death of Jesus is one of the events of redemption, but his death alone is not the totality of salvation. His death is sacrificial, but his death is not the complete sacrifice for salvation without his resurrection.", False, False)
    ])

    p_body([
        ("There are three cardinal events that mark the gospel of Christ, as articulated in 1 Corinthians 15:", False, False)
    ])

    script_ref("1 Corinthians 15:3-4")
    script_verse("[3] For I delivered unto you first of all that which I also received, how that Christ died for our sins according to the scriptures;")
    script_verse("[4] And that he was buried, and that he rose again the third day according to the scriptures.")

    p_body([
        ("The gospel is not merely that Christ died; the gospel is that Christ died, he was buried, and he rose again from the dead according to the scriptures. Notice Paul's emphatic reasoning in verses 14 and 17:", False, False)
    ])

    script_ref("1 Corinthians 15:14, 17")
    script_verse("[14] And if Christ be not risen, then is our preaching vain, and your faith is also vain.")
    script_verse("[17] And if Christ be not raised, your faith is vain; ye are yet in your sins.")

    p_body([
        ("Paul did not say, 'If Christ did not die, you are yet in your sins.' He stated plainly: if Christ was not raised from the dead, your faith is vain and you are still in your sins. Therefore, the resurrection of Jesus from the dead is the crown of redemption. It is the divine proof that redemption has been accomplished and that man has been justified. If Christ were not raised, the penalty would not be cleared; his resurrection confirms the total remission of sins.", False, False)
    ])

    # --- SECTION 3 ---
    h2("3. The New Testament Language of \"IN\": Ezekiel vs. Joel and the Pauline Reality")

    p_body([
        ("In the Gospel of John, Jesus expounds upon the nature of the new birth. When speaking to Nicodemus in John 3:", False, False)
    ])

    script_ref("John 3:3, 5-8")
    script_verse("[3] Jesus answered and said unto him, Verily, verily, I say unto thee, Except a man be born again, he cannot see the kingdom of God.")
    script_verse("[5] Jesus answered, Verily, verily, I say unto thee, Except a man be born of water and of the Spirit, he cannot enter into the kingdom of God.")
    script_verse("[6] That which is born of the flesh is flesh; and that which is born of the Spirit is spirit.")
    script_verse("[7] Marvel not that I said unto thee, Ye must be born again.")
    script_verse("[8] The wind bloweth where it listeth, and thou hearest the sound thereof, but canst not tell whence it cometh, and whither it goeth: so is every one that is born of the Spirit.")

    p_body([
        ("Jesus uses the phrase 'born of the Spirit' three times. Being 'born of the Spirit' explains what it means to be 'born again.' An individual does not give birth to himself; the parent determines the birth. In salvation, the believer is born of the Spirit of God. In John 4:14, Jesus describes this indwelling life as an interior fountain:", False, False)
    ])

    script_ref("John 4:14")
    script_verse("[14] But whosoever drinketh of the water that I shall give him shall never thirst; but the water that I shall give him shall be in him a well of water springing up into everlasting life.")

    p_body([
        ("Likewise, in John 7:37-39, the Spirit is spoken of as living water received within the believer:", False, False)
    ])

    script_ref("John 7:37-39")
    script_verse("[37] In the last day, that great day of the feast, Jesus stood and cried, saying, If any man thirst, let him come unto me, and drink.")
    script_verse("[38] He that believeth on me, as the scripture hath said, out of his belly shall flow rivers of living water.")
    script_verse("[39] (But this spake he of the Spirit, which they that believe on him should receive...)")

    p_body([
        ("In salvation, something divine is given and received. In John 14:16-17, Jesus identifies this indwelling presence:", False, False)
    ])

    script_ref("John 14:16-17")
    script_verse("[16] And I will pray the Father, and he shall give you another Comforter, that he may abide with you for ever;")
    script_verse("[17] Even the Spirit of truth; whom the world cannot receive, because it seeth him not, neither knoweth him: but ye know him; for he dwelleth with you, and shall be in you.")

    p_body([
        ("The hallmark language of the New Testament is the preposition ", False, False),
        ("\"IN\"", True, False),
        (". It is 'in Christ' and 'in you.' In the Old Testament prophecies, Joel prophesied that the Spirit would be poured ", False, False),
        ("upon", True, True),
        (" all flesh (Joel 2:28), whereas Ezekiel prophesied that God would put His Spirit ", False, False),
        ("within", True, True),
        (" man (Ezekiel 36:27). Ezekiel's prophecy is closer to the New Testament reality because the definitive language of the new covenant is the Spirit dwelling ", False, False),
        ("within", True, False),
        (". As Paul wrote to Philemon, the communication of faith becomes effectual by acknowledging every good thing which is ", False, False),
        ("in you in Christ Jesus", True, False),
        (" (Philemon 6).", False, False)
    ])

    p_body([
        ("While Luke records the historical commission in Luke 24:49 ('until ye be endued with power from on high') and Acts 1:8 ('power, after that the Holy Ghost is come upon you'), John's Gospel presents the internal, ontological reality: the Spirit abides ", False, False),
        ("in you forever", True, False),
        (". Paul reinforces this systematically in Romans 8:", False, False)
    ])

    script_ref("Romans 8:1-2, 9-11")
    script_verse("[1] There is therefore now no condemnation to them which are in Christ Jesus...")
    script_verse("[2] For the law of the Spirit of life in Christ Jesus hath made me free from the law of sin and death.")
    script_verse("[9] But ye are not in the flesh, but in the Spirit, if so be that the Spirit of God dwell in you. Now if any man have not the Spirit of Christ, he is none of his.")
    script_verse("[10] And if Christ be in you, the body is dead because of sin; but the Spirit is life because of righteousness.")
    script_verse("[11] But if the Spirit of him that raised up Jesus from the dead dwell in you, he that raised up Christ from the dead shall also quicken your mortal bodies by his Spirit that dwelleth in you.")

    p_body([
        ("Notice the repetitive usage: 'in Christ', 'in us', 'in you.' To be in Christ and to have the Spirit dwelling in you are synonymous truths. If an individual has not the Spirit of Christ, he does not belong to him. When an individual has the Spirit of Christ, he is in Christ and Christ is in him. In Romans 8:12-14, the statement 'if ye live after the flesh, ye shall die' presents an impossibility for the regenerate believer, because verse 14 clarifies: 'For as many as are led by the Spirit of God, they are the sons of God.' The believer has received the Spirit of adoption, witnessing with his spirit that he is a child of God.", False, False)
    ])

    # --- SECTION 4 ---
    h2("4. The Believer's Total Redemption: Refuting the Dual Personality and Generational Curse Theories")

    p_body([
        ("Because salvation is the reception of the Spirit of God, popular theories proposing that a Christian possesses dual personalities or houses demonic spirits alongside the Holy Spirit represent completely erroneous theology. Some teachers propose a flawed maxim: 'You can take the man out of the dirt in salvation, but it takes sanctification and deliverance to take the dirt out of the man.' Under this erroneous doctrine, believers are told that foreign spirits remain in their lineage or bodies, requiring ritual deliverance from water spirits, family shrines, or ancestral curses.", False, False)
    ])

    p_anecdote(
        "Early in my Christian walk, I fell ill and sought counseling. The counselor inquired extensively into my family background. When I shared that my ancestors had worshipped at shrines, he shook his head and asserted that water goddesses were following me, claiming my illness was due to ancestral claims. He instructed me to go home and ask my mother what transpired at my birth, telling me that we had to appease or beg these deities to take their hands off me."
    )

    p_body([
        ("Such reasoning is spiritually bankrupt. Consider Abraham: Abraham was called out of an idol-worshipping home in Ur of the Chaldees (Joshua 24:2). Terah, Abraham's father, served other gods. Yet God did not administer ritual deliverance or investigate ancestral water spirits before establishing His covenant with Abraham. In Pauline theology, sin passed upon all men through Adam (Romans 5:12). Every human being was born under the universal dominion of sin, and every believer is completely redeemed through Christ. An idol or a mythical goddess cannot give birth to a human child; humans produce humans. To suggest that a believer remains tied to demonic parentage denies the total efficacy of the new birth.", False, False)
    ])

    p_body([
        ("Salvation is not a product of subjective emotional experiences or ancestral background; it is a gift from God. The Holy Spirit does not wait for a vessel to become holy before entering; rather, the Holy Spirit enters the sinner the moment he believes the gospel, and His indwelling presence is what makes the believer holy. God gives His Holy Spirit to sinners who believe, cleansing them entirely.", False, False)
    ])

    # --- SECTION 5 ---
    h2("5. The Holy Spirit as the Spiritual DNA of the Believer: Sealed for Permanent Ownership")

    p_body([
        ("Paul details the irrevocable nature of this indwelling in Ephesians 1 and 4:", False, False)
    ])

    script_ref("Ephesians 1:13 & 4:30")
    script_verse("[1:13] In whom ye also trusted, after that ye heard the word of truth, the gospel of your salvation: in whom also after that ye believed, ye were sealed with that holy Spirit of promise.")
    script_verse("[4:30] And grieve not the holy Spirit of God, whereby ye are sealed unto the day of redemption.")

    p_body([
        ("In the ancient Roman Empire, a seal (*sphragis*) was a legal signet representing two inviolable realities: proof of ownership and permanent security. When God placed His Holy Spirit within the believer, He sealed him as His permanent possession. God claims the believer totally—his past, his present, and his future.", False, False)
    ])

    p_body([
        ("The Holy Spirit is therefore the spiritual DNA of the Christian. Just as a biological child carries the indelible genetics and heredity of his parents—which can never be altered or disclaimed in a court of law, regardless of personal disowning—the believer carries within him the divine nature of God. The Spirit of God is not an intermittent experience; He is the very life received from God. As 1 Corinthians 2:12 affirms, we have received the Spirit which is of God, that we might know the things that are freely given to us of God.", False, False)
    ])

    # --- SECTION 6 ---
    h2("6. Understanding \"Greater Works\": Refuting the Cessationist View of Charles Spurgeon's Successor")

    p_body([
        ("In John 14:12, Jesus makes a pivotal statement regarding the ministry of the believer:", False, False)
    ])

    script_ref("John 14:12")
    script_verse("[12] Verily, verily, I say unto you, He that believeth on me, the works that I do shall he do also; and greater works than these shall he do; because I go unto my Father.")

    p_body([
        ("Throughout church history, this verse has often been subjected to horrific interpretative gymnastics by theologians attempting to explain away the miraculous.", False, False)
    ])

    p_anecdote(
        "A prominent theologian who pastored the Metropolitan Tabernacle in London—the church originally founded and pastored by Charles Spurgeon—authored a book addressing John 14:12. In his writing, he asserted that Jesus was indicating that the days of physical miracles would cease. He argued that 'greater works' could only refer to salvation, asking rhetorically: 'What can be greater than what Jesus did? Winning a soul is greater than healing the sick, so Jesus meant the winning of souls.' When a reader shared this commentary with me, marveling at the explanation, I pointed out the obvious textual error."
    )

    p_body([
        ("Jesus did not say, 'He that believeth on me will no longer be able to do the works that I do.' Jesus said explicitly: 'the works that I do shall he do ", False, False),
        ("ALSO", True, False),
        (".' The word ", False, False),
        ("'also'", True, False),
        (" means the believer will heal the sick, cleanse the lepers, raise the dead, and cast out demons just as Jesus did. Furthermore, Jesus said ", False, False),
        ("'greater works' (plural, Greek: erga)", True, False),
        (", signifying deeds and miraculous acts, not a singular concept. Moreover, salvation is not a work that man performs; salvation is the work of God. One cannot substitute human soul-winning as a 'greater work' to nullify the miraculous power of God in the believer's hands.", False, False)
    ])

    # --- SECTION 7 ---
    h2("7. The Terminology of Spiritual Gifts: Dorea and Utterance Gifts in the Book of Acts")

    p_body([
        ("In the four Gospels, Jesus introduced the truths of the Spirit, but the full apostolic exposition was delivered after his ascension. As Jesus declared in John 16:12-13, the disciples could not bear the fullness of truth until the Spirit of truth arrived to guide them into all truth. In Mark 16:17-18, Jesus declared the signs that accompany those who believe: casting out demons, speaking with new tongues, taking up serpents, drinking deadly things without harm, and laying hands on the sick for recovery.", False, False)
    ])

    p_body([
        ("In the Book of Acts, the term used for 'gift' is the Greek word ", False, False),
        ("***dorea*** (or *doreon*)", True, False),
        (", meaning a free, unmerited gift or offering. It appears in Peter's sermon in Acts 2:38:", False, False)
    ])

    script_ref("Acts 2:38")
    script_verse("[38] Then Peter said unto them, Repent, and be baptized every one of you in the name of Jesus Christ for the remission of sins, and ye shall receive the gift of the Holy Ghost.")

    p_body([
        ("It appears again in Acts 8:20, when Peter rebukes Simon the sorcerer for thinking that 'the gift of God (*dorea*) may be purchased with money.' An examination of the Book of Acts reveals that whenever the phrase 'the gift of the Holy Ghost' or the preposition 'upon' (*epi*) is used, it refers specifically to the ", False, False),
        ("utterance gifts", True, False),
        ("—speaking in tongues and prophesying.", False, False)
    ])

    p_body([
        ("In Acts 2:16-17, Peter quotes Joel's prophecy: 'I will pour out of my Spirit upon all flesh: and your sons and your daughters shall prophesy.' In Acts 8:15-16, the Holy Ghost was 'fallen upon none of them' until Peter and John laid hands on them, at which point tangible utterance manifested. In Acts 10:44-46, as Peter spoke, 'the Holy Ghost fell on all them which heard the word... for they heard them speak with tongues, and magnify God.' In Acts 11:17, Peter reflects on this event, describing it as 'the like gift' (*dorea*). In Acts 19:6, when Paul laid hands on the disciples at Ephesus, 'the Holy Ghost came upon them; and they spake with tongues, and prophesied.' Thus, in the historical narrative of Acts, the term *dorea* and the Spirit coming 'upon' consistently designate utterance.", False, False)
    ])

    # --- SECTION 8 ---
    h2("8. Revelation Gifts Expressed through Utterance: Ananias, Peter, and Agabus")

    p_body([
        ("Similarly, what the church recognizes as the revelation gifts (word of wisdom, word of knowledge, discerning of spirits) are described in the Book of Acts using the language of the Spirit speaking or signifying:", False, False)
    ])

    script_ref("Acts 5:3-4, 9")
    script_verse("[3] But Peter said, Ananias, why hath Satan filled thine heart to lie to the Holy Ghost, and to keep back part of the price of the land?")
    script_verse("[4] ...thou hast not lied unto men, but unto God.")
    script_verse("[9] Then Peter said unto her, How is it that ye have agreed together to tempt the Spirit of the Lord?")

    p_body([
        ("Peter knew the deception of Ananias and Sapphira by the revelation of the Spirit, and he spoke it out. In Acts 10:19, while Peter thought on the vision on the rooftop, 'the Spirit said unto him, Behold, three men seek thee.' In Acts 11:28, the prophet Agabus stood up and 'signified by the Spirit that there should be great dearth throughout all the world.' In Acts 13:2, as the prophets and teachers ministered to the Lord and fasted, 'the Holy Ghost said, Separate me Barnabas and Saul for the work whereunto I have called them.' In Acts 20:23, Paul testified that 'the Holy Ghost witnesseth in every city, saying that bonds and afflictions abide me.' In Acts 21:11, Agabus took Paul's girdle, bound himself, and proclaimed: 'Thus saith the Holy Ghost, So shall the Jews at Jerusalem bind the man that owneth this girdle.'", False, False)
    ])

    p_body([
        ("Notice the consistent operational reality: revelation gifts are communicated through utterance. Agabus spoke out the revelation; Peter spoke out the revelation; the prophets at Antioch spoke out the direction of the Spirit. An individual who walks in the utterance gifts possesses the scriptural platform to speak forth revelation gifts.", False, False)
    ])

    # --- SECTION 9 ---
    h2("9. Charisma as Spiritual Endowment: The Natural and Supernatural Dimensions of Ministry")

    p_body([
        ("When Paul addresses the local assembly in 1 Corinthians 12, he transitions from narrative description to theological precision:", False, False)
    ])

    script_ref("1 Corinthians 12:1, 4")
    script_verse("[1] Now concerning spiritual gifts, brethren, I would not have you ignorant.")
    script_verse("[4] Now there are diversities of gifts, but the same Spirit.")

    p_body([
        ("In verse 1, the word translated 'spiritual gifts' is ", False, False),
        ("***pneumatikos*** (or *pneumatika*)", True, False),
        (", meaning 'spiritual things' or matters pertaining to the Spirit. In verse 4, Paul introduces the word ", False, False),
        ("***charisma***", True, False),
        (". *Charisma* means an endowment or an ability. In classic Greek literature, philosophers such as Plato and Aristotle used *charisma* to describe natural talents, artistic endowments, and inherent proficiencies in speech or craft.", False, False)
    ])

    p_body([
        ("Paul adopts this term to explain that the manifestations of the Spirit are supernatural abilities and endowments that have become the very nature of the believer through the indwelling Spirit. Just as in natural athletics, a gifted football player like Lionel Messi or Cristiano Ronaldo possesses an innate endowment for the game that makes extraordinary performance look effortless, or a versatile actor like Arnold Schwarzenegger possesses the endowment to play distinctly different roles (transitioning from the antagonist in *Terminator 1* to the protector in *Terminator 2*), the believer is endowed by the Spirit with spiritual abilities that produce supernatural results.", False, False)
    ])

    p_body([
        ("Following the classic classification articulated by Bible teacher Howard Carter, the nine gifts of the Spirit fall into three distinct categories:", False, False)
    ])

    p_body([
        ("1. ", True, False),
        ("Utterance Gifts (Gifts that say something):", True, False),
        (" Tongues, Interpretation of Tongues, Prophecy.\n", False, False),
        ("2. ", True, False),
        ("Revelation Gifts (Gifts that show something):", True, False),
        (" Word of Wisdom, Word of Knowledge, Discerning of Spirits.\n", False, False),
        ("3. ", True, False),
        ("Power Gifts (Gifts that do something):", True, False),
        (" Special Faith, Working of Miracles, Gifts of Healings.", False, False)
    ])

    p_body([
        ("The crucial truth is that these manifestations are not passive, uncontrollable lightning bolts; they are endowments—abilities placed within the believer by the Spirit that produce supernatural experiences.", False, False)
    ])

    # --- SECTION 10 ---
    h2("10. The Gifts of Healings: Laying on of Hands as a Deliberate Touch for Ministering Power")

    p_body([
        ("In 1 Corinthians 12:9, 28, and 30, the apostle Paul specifies the healing ministry using a unique plural construction in the Greek text: ", False, False),
        ("***charismata iamaton*** — \"gifts of healings\"", True, False),
        (". It is plural in both words: not 'the gift of healing,' but 'gifts of healings.'", False, False)
    ])

    script_ref("1 Corinthians 12:9, 28, 30")
    script_verse("[9] To another faith by the same Spirit; to another the gifts of healing by the same Spirit;")
    script_verse("[28] And God hath set some in the church, first apostles, secondarily prophets, thirdly teachers, after that miracles, then gifts of healings, helps, governments, diversities of tongues.")
    script_verse("[30] Have all the gifts of healing? do all speak with tongues? do all interpret?")

    p_body([
        ("The plural form 'gifts of healings' demonstrates that the endowment to minister healing is characterized by diverse administrations and methods. It does not refer to different gifts for different diseases (e.g., one gift for malaria and another for cancer); rather, it refers to diverse ways in which divine healing power is ministered to the sick.", False, False)
    ])

    p_body([
        ("The foundational and primary administration of the gifts of healings is ", False, False),
        ("the laying on of hands", True, False),
        (". Certain modern teachers exhibit profound ignorance by teaching that laying on of hands is an inferior or 'baby' method of receiving healing, claiming that 'speaking the word' is a superior method. Scripture knows no such hierarchy. Jesus himself commanded in the Great Commission:", False, False)
    ])

    script_ref("Mark 16:18")
    script_verse("[18] ...they shall lay hands on the sick, and they shall recover.")

    p_body([
        ("In the Greek language, 'laying on of hands' (*epitithemi cheiras*) means a ", False, False),
        ("deliberate, purposeful physical touch", True, False),
        (" intended to minister and transmit the power of God into the body of the sick individual.", False, False)
    ])

    # --- SECTION 11 ---
    h2("11. Scriptural Dynamics of Laying on of Hands: Healing the Sick vs. Ecclesiastical Ordination")

    p_body([
        ("It is vital to distinguish between the two primary scriptural contexts where laying on of hands is practiced: healing the sick and ecclesiastical ordination.", False, False)
    ])

    p_body([
        ("In 1 Timothy 5:22, Paul writes: 'Lay hands suddenly on no man, neither be partaker of other men's sins: keep thyself pure.' Some individuals mistakenly apply this verse to healing services, arguing that ministering to the sick by laying on of hands will cause the minister to contract the sins or sicknesses of the person. That is complete rubbish. In 1 Timothy 5, Paul is addressing ecclesiastical ordination—appointing individuals to elder or pastoral oversight. The Greek word for 'suddenly' is ", False, False),
        ("***tacheos***", True, False),
        (", meaning hastily or rashly. Paul is instructing Timothy not to ordain a novice into public ministry without proper vetting and proving. In Acts 13:3, the church laid hands on Barnabas and Saul to commission them for their apostolic journey. Ordination is an act of commissioning and appointment.", False, False)
    ])

    p_body([
        ("In the realm of healing, however, laying on of hands is a deliberate physical touch upon the sick to impart divine virtue. Consider the ministry of Jesus in the Gospels:", False, False)
    ])

    script_ref("Matthew 8:2-3, 14-15")
    script_verse("[2] And, behold, there came a leper and worshipped him, saying, Lord, if thou wilt, thou canst make me clean.")
    script_verse("[3] And Jesus put forth his hand, and touched him, saying, I will; be thou clean. And immediately his leprosy was cleansed.")
    script_verse("[14] And when Jesus was come into Peter's house, he saw his wife's mother laid, and sick of a fever.")
    script_verse("[15] And he touched her hand, and the fever left her: and she arose, and ministered unto them.")

    p_body([
        ("Notice that Jesus put forth his hand and ", False, False),
        ("touched", True, False),
        (" the leper. In the case of Peter's mother-in-law (revealing humorously that the first supposed 'Pope' had a wife!), Jesus touched her hand, and the fever broke immediately. In Mark 8:22-25, when a blind man was brought to Jesus at Bethsaida, Jesus took him by the hand, spat on his eyes, and put his hands upon him; when the man saw partially, Jesus put his hands again upon his eyes until his sight was restored completely. In Mark 10:16, Jesus took little children up in his arms, put his hands upon them, and blessed them.", False, False)
    ])

    p_body([
        ("Most remarkably, in Mark 6:5, in his hometown of Nazareth where there was widespread offense and unbelief:", False, False)
    ])

    script_ref("Mark 6:5")
    script_verse("[5] And he could there do no mighty work, save that he laid his hands upon a few sick folk, and healed them.")

    p_body([
        ("Even where atmospheric unbelief hindered public corporate miracles, the deliberate laying on of hands upon sick individuals still broke through and healed them. The touch of faith transmitted divine virtue.", False, False)
    ])

    # --- SECTION 12 ---
    h2("12. Power in Action: Diverse Administrations and the Transmission of Divine Healing")

    p_body([
        ("The same practice continued throughout the apostolic era in the Book of Acts:", False, False)
    ])

    script_ref("Acts 28:8-9")
    script_verse("[8] And it came to pass, that the father of Publius lay sick of a fever and of a bloody flux: to whom Paul entered in, and prayed, and laid his hands on him, and healed him.")
    script_verse("[9] So when this was done, others also, which had diseases in the island, came, and were healed.")

    p_body([
        ("On the island of Malta, Paul entered the quarters of the chief official whose father was dying of dysentery and fever. Paul prayed, laid his hands upon him, and healed him. Once that was done, all the other diseased islanders came and were healed. In Acts 3:6-7, at the Gate Beautiful, Peter took the lame man by the right hand and lifted him up; through that deliberate physical contact, strength entered his ankle bones. In Acts 8:5-7, Philip preached Christ in Samaria, and the multitudes gave heed when they heard and saw the miracles which he did in action.", False, False)
    ])

    p_body([
        ("Laying on of hands is not an empty formal ritual of placing hands randomly upon someone's head. When someone suffers from leg pain, placing hands upon the head misses the targeted point of affliction. A deliberate touch involves placing hands with modesty, decency, and spiritual authority directly where the pain or infirmity resides, deliberately releasing the power of God. It is a purposeful physical contact infused with the power of the Holy Spirit.", False, False)
    ])

    p_body([
        ("The power of God is in the action. The believer does not merely wish or hope; he acts upon the authority of Christ. When hands are laid upon the sick in faith, the divine life and virtue resident within the believer's spirit are transmitted into the physical body of the afflicted. As believers develop understanding and skill in these power gifts, sickness is arrested, divine health is ministered, and the Lord Jesus is magnified across the earth.", False, False)
    ])

    # --- EXPORT ---
    for docx_p in docx_paths:
        Path(docx_p).parent.mkdir(parents=True, exist_ok=True)
        doc.save(docx_p)
        print(f"[+] Saved Word docx to: {docx_p}")

    full_md = "\n".join(md_lines)
    for md_p in md_paths:
        Path(md_p).parent.mkdir(parents=True, exist_ok=True)
        with open(md_p, "w", encoding="utf-8") as f:
            f.write(full_md)
        print(f"[+] Saved Markdown to: {md_p}")

    # Raw transcript is managed independently as verbatim oral audio transcript
    if raw_path and not Path(raw_path).exists():
        print(f"[!] Note: Verbatim raw transcript should be placed at: {raw_path}")

if __name__ == '__main__':
    base_dir = Path(r"c:\Users\Stephycopy\OneDrive\Desktop\Theos\theos")
    spirit_dir = base_dir / "In the Spirit by the Spirit Transcript"

    docx_files = [
        base_dir / "In the Spirit by the Spirit – Power Gifts - Track 2.docx",
        spirit_dir / "Formatted" / "In the Spirit by the Spirit – Power Gifts - Track 2.docx"
    ]

    md_files = [
        base_dir / "In the Spirit by the Spirit – Power Gifts - Track 2.md",
        spirit_dir / "Formatted" / "In the Spirit by the Spirit – Power Gifts - Track 2.md"
    ]

    raw_file = spirit_dir / "Raw" / "In the Spirit by the Spirit – Power Gifts - Track 2_Raw.txt"

    build_power_gifts_transcript(docx_files, md_files, raw_file)

    doc = Document(docx_files[0])
    total_words = sum(len(p.text.split()) for p in doc.paragraphs)
    print(f"[+] Verified: Total word count in Word document = {total_words} words.")
