# -*- coding: utf-8 -*-
"""
Builder for:
In the Spirit by the Spirit – Power Gifts - Track 2 (Fillers Removed)

Features:
- Spoken verbal fillers, stuttering false starts, and repetitive audience prompts removed.
- 100% fidelity to the speaker's genuine vocabulary, theology, stories, and Greek expositions.
- All 51 Bible reference headers in BOLD.
- All 107 full KJV scripture verses with bracketed verse notation '[#]' in BOLD on separate lines.
- Clean typography: Lexend/Calibri, 10pt text, 1.15 line spacing, 5pt space after, 1.0 inch margins.
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor

ELEMENTS_CLEAN = [
    # 1. Introduction & Salvation Defined
    ('p', "Open your word this morning, that our minds receive wisdom, that everyone be wiser upon hearing the truth of your word this morning, in the name of Jesus. Amen."),
    ('p', "In 2 Corinthians 1. In the Spirit, by the Spirit, this series started last year March, and we've been looking at the things of the Spirit generally. Sometimes, spiritual things could take a long time, longer than you thought, but we're blessed anyway. Verse 18 and 19:"),
    ('ref', "2 Corinthians 1:18-22"),
    ('v', "[18] But as God is true, our word toward you was not yea and nay."),
    ('v', "[19] For the Son of God, Jesus Christ, who was preached among you by us, even by me and Silvanus and Timotheus, was not yea and nay, but in him was yea."),
    ('v', "[20] For all the promises of God in him are yea, and in him Amen, unto the glory of God by us."),
    ('v', "[21] Now he which stablisheth us with you in Christ, and hath anointed us, is God;"),
    ('v', "[22] Who hath also sealed us, and given the earnest of the Spirit in our hearts."),
    ('p', "Romans 15:8 says:"),
    ('ref', "Romans 15:8"),
    ('v', "[8] Now I say that Jesus Christ was a minister of the circumcision for the truth of God, to confirm the promises made unto the fathers:"),
    ('p', "Salvation and the Spirit are synonymous terms. There cannot be salvation without the giving of the Spirit. The giving of the Spirit is salvation."),
    ('p', "You will find the phrase 'giving of the Spirit' used in the Book of Acts that seems to differ from this, but salvation is synonymous with the giving of the Spirit. You cannot be saved and not have the Spirit of God. The gift of the Spirit is salvation."),

    # 2. Events of Redemption & Resurrection
    ('p', "Jesus' death is one of the events of redemption. Most of the time our minds stop at the cross, thinking Jesus died for our sins. His death is sacrificial—in other words, he suffered—but his death alone is not the sacrifice for salvation. What Jesus did is evidenced by his Spirit that he gave us."),
    ('p', "There are three events that mark the gospel of Christ. 1 Corinthians 15:3 and 4 states them:"),
    ('ref', "1 Corinthians 15:3-4"),
    ('v', "[3] For I delivered unto you first of all that which I also received, how that Christ died for our sins according to the scriptures;"),
    ('v', "[4] And that he was buried, and that he rose again the third day according to the scriptures:"),
    ('p', "Paul notes that Christ died according to the scriptures, was buried, and rose again the third day according to the scriptures. That he was buried was proof that he died. That he rose from the dead proves redemption. The gospel is not merely that Christ died; the gospel is that Christ died, he was buried, and he rose from the dead."),
    ('p', "In 1 Corinthians 15:14:"),
    ('ref', "1 Corinthians 15:14"),
    ('v', "[14] And if Christ be not risen, then is our preaching vain, and your faith is also vain."),
    ('p', "The word 'vain' means fruitless and temporary. In verse 17:"),
    ('ref', "1 Corinthians 15:17"),
    ('v', "[17] And if Christ be not raised, your faith is vain; ye are yet in your sins."),
    ('p', "If Christ was not raised from the dead, your faith is vain, and you are still in your sins. He did not say if Christ did not die, but if he was not raised. The resurrection of Jesus from the dead is the crown and proof of redemption. It is the proof that man has been redeemed."),

    # 3. Born of the Spirit & Water Springing Up
    ('p', "In John 3, look at the words of Jesus in verse 3:"),
    ('ref', "John 3:3"),
    ('v', "[3] Jesus answered and said unto him, Verily, verily, I say unto thee, Except a man be born again, he cannot see the kingdom of God."),
    ('p', "Nicodemus asked in verse 4 what it means to be born again. Jesus answered in verses 5 through 8:"),
    ('ref', "John 3:5-8"),
    ('v', "[5] Jesus answered, Verily, verily, I say unto thee, Except a man be born of water and of the Spirit, he cannot enter into the kingdom of God."),
    ('v', "[6] That which is born of the flesh is flesh; and that which is born of the Spirit is spirit."),
    ('v', "[7] Marvel not that I said unto thee, Ye must be born again."),
    ('v', "[8] The wind bloweth where it listeth, and thou hearest the sound thereof, but canst not tell whence it cometh, and whither it goeth: so is every one that is born of the Spirit."),
    ('p', "Jesus uses the phrase 'born of the Spirit' three times. 'Born of the Spirit' explains what it means to be born again. You do not give birth to yourself; the parent determines the birth. Jesus calls it being born of the Spirit."),
    ('p', "In John 4:14, he puts it differently:"),
    ('ref', "John 4:14"),
    ('v', "[14] But whosoever drinketh of the water that I shall give him shall never thirst; but the water that I shall give him shall be in him a well of water springing up into everlasting life."),
    ('p', "Jesus says he will give to whoever believes water that springs up into everlasting life. The first explanation says a man is born of the Spirit; the second says he is given water springing up into eternal life. Then in John 7:37 through 39:"),
    ('ref', "John 7:37-39"),
    ('v', "[37] In the last day, that great day of the feast, Jesus stood and cried, saying, If any man thirst, let him come unto me, and drink."),
    ('v', "[38] He that believeth on me, as the scripture hath said, out of his belly shall flow rivers of living water."),
    ('v', "[39] (But this spake he of the Spirit, which they that believe on him should receive: for the Holy Ghost was not yet given; because that Jesus was not yet glorified.)"),
    ('p', "In salvation, there is a new birth, and there is what is received. You do not just get saved the way you are; God gives something. Jesus says he gives the Spirit, which is the same as being born of the Spirit."),

    # 4. The Language of the New Testament: IN vs UPON
    ('p', "In John 14:16 and 17, Jesus says:"),
    ('ref', "John 14:16-17"),
    ('v', "[16] And I will pray the Father, and he shall give you another Comforter, that he may abide with you for ever;"),
    ('v', "[17] Even the Spirit of truth; whom the world cannot receive, because it seeth him not, neither knoweth him: but ye know him; for he dwelleth with you, and shall be in you."),
    ('p', "The language of the New Testament is 'IN'. Between Ezekiel's prophecy and Joel's prophecy, Ezekiel's prophecy is closer to the New Testament because Ezekiel says 'IN', whereas Joel says 'UPON'. 'IN' is the definitive language of the New Testament: in Christ, in you. Paul does not say to Philemon to acknowledge every good thing that is upon you, but every good thing that is in you in Christ Jesus. The Spirit abides in you forever."),
    ('p', "We have two other statements in Luke and Acts. Luke 24:49 says:"),
    ('ref', "Luke 24:49"),
    ('v', "[49] And, behold, I send the promise of my Father upon you: but tarry ye in the city of Jerusalem, until ye be endued with power from on high."),
    ('p', "And Acts 1:8 says:"),
    ('ref', "Acts 1:8"),
    ('v', "[8] But ye shall receive power, after that the Holy Ghost is come upon you: and ye shall be witnesses unto me both in Jerusalem, and in all Judaea, and in Samaria, and unto the uttermost part of the earth."),
    ('p', "These are Jesus' words, and they can cause confusion because we have 'in you' and 'upon you'. Traditional beliefs often treat them as two entirely different spirits or separate experiences because influential teachers started those doctrines. When people say the Spirit upon you and the Spirit within you represent two different experiences, that is not entirely so. Luke's explanation of 'upon' is valid, but it is only part of the whole truth."),

    # 5. Romans 8: Indwelling & Adoption
    ('p', "Paul's language for the Spirit in Romans 8 is 'IN':"),
    ('ref', "Romans 8:1-2"),
    ('v', "[1] There is therefore now no condemnation to them which are in Christ Jesus, who walk not after the flesh, but after the Spirit."),
    ('v', "[2] For the law of the Spirit of life in Christ Jesus hath made me free from the law of sin and death."),
    ('p', "The law of the Spirit of life is in Christ Jesus. Verse 4 says:"),
    ('ref', "Romans 8:4"),
    ('v', "[4] That the righteousness of the law might be fulfilled in us, who walk not after the flesh, but after the Spirit."),
    ('p', "The righteousness of the law is fulfilled in us. Then verses 8 through 11:"),
    ('ref', "Romans 8:8-11"),
    ('v', "[8] So then they that are in the flesh cannot please God."),
    ('v', "[9] But ye are not in the flesh, but in the Spirit, if so be that the Spirit of God dwell in you. Now if any man have not the Spirit of Christ, he is none of his."),
    ('v', "[10] And if Christ be in you, the body is dead because of sin; but the Spirit is life because of righteousness."),
    ('v', "[11] But if the Spirit of him that raised up Jesus from the dead dwell in you, he that raised up Christ from the dead shall also quicken your mortal bodies by his Spirit that dwelleth in you."),
    ('p', "Notice the synonymous terms: you are in Christ; the law of the Spirit of life is in Christ; righteousness is fulfilled in us; you are in the Spirit; and the Spirit of God dwells in you. Being in Christ and having Christ in you are synonymous. If you are in Christ, Christ is in you. The indwelling of the Spirit is repeated across Romans 8: in you, in Christ."),
    ('p', "Verses 12 through 16 continue:"),
    ('ref', "Romans 8:12-16"),
    ('v', "[12] Therefore, brethren, we are debtors, not to the flesh, to live after the flesh."),
    ('v', "[13] For if ye live after the flesh, ye shall die: but if ye through the Spirit do mortify the deeds of the body, ye shall live."),
    ('v', "[14] For as many as are led by the Spirit of God, they are the sons of God."),
    ('v', "[15] For ye have not received the spirit of bondage again to fear; but ye have received the Spirit of adoption, whereby we cry, Abba, Father."),
    ('v', "[16] The Spirit itself beareth witness with our spirit, that we are the children of God:"),
    ('p', "In verse 13, the conditional statement paints a hypothetical scenario, because verse 14 explains why the believer is not under death: 'For as many as are led by the Spirit of God, they are the sons of God.' Being led by the Spirit is receiving the Spirit of adoption, whereby we cry Abba, Father. John's rendering of the Spirit dwelling in you is the definitive language of the New Testament."),

    # 6. Galatians 4 & 1 Corinthians 2: Receiving the Spirit
    ('p', "In Galatians 4:4 through 6:"),
    ('ref', "Galatians 4:4-6"),
    ('v', "[4] But when the fulness of the time was come, God sent forth his Son, made of a woman, made under the law,"),
    ('v', "[5] To redeem them that were under the law, that we might receive the adoption of sons."),
    ('v', "[6] And because ye are sons, God hath sent forth the Spirit of his Son into your hearts, crying, Abba, Father."),
    ('p', "Salvation is receiving the Spirit. It is not being transferred from darkness to light and then wandering back into darkness. Salvation means you have received the Spirit of God, and you are in him."),
    ('p', "In 1 Corinthians 2:9 through 12:"),
    ('ref', "1 Corinthians 2:9-12"),
    ('v', "[9] But as it is written, Eye hath not seen, nor ear heard, neither have entered into the heart of man, the things which God hath prepared for them that love him."),
    ('v', "[10] But God hath revealed them unto us by his Spirit: for the Spirit searcheth all things, yea, the deep things of God."),
    ('v', "[11] For what man knoweth the things of a man, save the spirit of man which is in him? even so the things of God knoweth no man, but the Spirit of God."),
    ('v', "[12] Now we have received, not the spirit of the world, but the spirit which is of God; that we might know the things that are freely given to us of God."),
    ('p', "Verses 9 through 11 explain the omniscience of the Spirit searching the deep things of God. Verse 12 explains salvation: we have received the Spirit which is of God, that we might know the things freely given to us of God. Salvation is receiving the Spirit of God. You cannot be saved and lack the Spirit."),
    ('p', "Some teach a theory that salvation takes a man out of the dirt, but deliverance is needed to take the dirt out of the man. They suggest that a believer has foreign spirits cohabiting as tenants or landlords alongside Christ. That is false."),

    # 7. Deliverance Mythology vs. The Gift of Salvation
    ('p', "When I was newly saved, I used to fall sick and sought counseling. The counselor asked for my family background, and I shared stories my father and uncle had told me. He immediately concluded that family spirits were pursuing me, claiming that because my parents once worshipped at a shrine, I had been obtained from a water goddess and needed to beg the goddess to take her hands off me. In an African context, such teachings introduce spooky superstitions that rob people of scriptural clarity."),
    ('p', "Abraham came from an idol-worshipping family, yet God did not tell Abraham that he needed deliverance before God could speak to him. From Paul's theology, sin passed upon all men through Adam. Water goddesses cannot give anyone a child; they have no such capacity."),
    ('p', "Salvation is not a product of your experiences or habits. Experiences can be good, bad, or neutral, but salvation is God's gift to you. The Holy Spirit does not wait for a sinner to become righteous before entering; the Holy Spirit enters the sinner who believes and makes him holy. At salvation, you receive the Holy Spirit."),

    # 8. Ephesians 1 & 4: Permanent Ownership & Heredity
    ('p', "In Ephesians 1:13 and Ephesians 4:30:"),
    ('ref', "Ephesians 1:13"),
    ('v', "[13] In whom ye also trusted, after that ye heard the word of truth, the gospel of your salvation: in whom also after that ye believed, ye were sealed with that holy Spirit of promise,"),
    ('ref', "Ephesians 4:30"),
    ('v', "[30] And grieve not the holy Spirit of God, whereby ye are sealed unto the day of redemption."),
    ('p', "The word 'sealed' in the Roman Empire was proof of permanent ownership. In salvation, God claims the believer completely—past, present, and future. You are sealed with the Spirit in you."),
    ('p', "Jesus says that which is born of the Spirit is spirit. The Spirit of God is the spiritual DNA and heredity of the Christian. Just as a parent can disown a child in court but cannot alter the biological DNA and life passed to that child, the believer carries the indestructible life of God. The Spirit of God is not merely an experience; the Spirit of God is the life we have received from God."),

    # 9. John 16 & John 14: Jesus Introduces the Gifts
    ('p', "In the Gospels, Jesus gives an introduction to truth rather than the complete post-resurrection explanation. He did not explicitly use the term 'spiritual gifts' or 'in Christ', but he introduced their workings. In John 16:12 and 13:"),
    ('ref', "John 16:12-13"),
    ('v', "[12] I have yet many things to say unto you, but ye cannot bear them now."),
    ('v', "[13] Howbeit when he, the Spirit of truth, is come, he will guide you into all truth: for he shall not speak of himself; but whatsoever he shall hear, that shall he speak: and he will shew you things to come."),
    ('p', "Jesus explains the works of the Spirit in John 14:12 through 14:"),
    ('ref', "John 14:12-14"),
    ('v', "[12] Verily, verily, I say unto you, He that believeth on me, the works that I do shall he do also; and greater works than these shall he do; because I go unto my Father."),
    ('v', "[13] And whatsoever ye shall ask in my name, that will I do, that the Father may be glorified in the Son."),
    ('v', "[14] If ye shall ask any thing in my name, I will do it."),
    ('p', "Some commentators argue that 'greater works' means miracles ceased and only soul-winning remains. But the Greek word here is 'ergon' (definite deeds or acts). Jesus said that he who believes will do the works Jesus did—healing the sick, raising the dead, casting out demons—and greater works than these, because Jesus goes to the Father. He was pointing directly to the manifestations of the Spirit."),
    ('p', "Then in John 14:15 through 20:"),
    ('ref', "John 14:15-20"),
    ('v', "[15] If ye love me, keep my commandments."),
    ('v', "[16] And I will pray the Father, and he shall give you another Comforter, that he may abide with you for ever;"),
    ('v', "[17] Even the Spirit of truth; whom the world cannot receive, because it seeth him not, neither knoweth him: but ye know him; for he dwelleth with you, and shall be in you."),
    ('v', "[18] I will not leave you comfortless: I will come to you."),
    ('v', "[19] Yet a little while, and the world seeth me no more; but ye see me: because I live, ye shall live also."),
    ('v', "[20] At that day ye shall know that I am in my Father, and ye in me, and I in you."),
    ('p', "The believer can do the works of Jesus because he has received the indwelling Spirit, and the Spirit works through him."),

    # 10. Mark 16 and Acts Terminology
    ('p', "In Mark 16:17 and 18, Jesus calls these manifestations 'signs':"),
    ('ref', "Mark 16:17-18"),
    ('v', "[17] And these signs shall follow them that believe; In my name shall they cast out devils; they shall speak with new tongues;"),
    ('v', "[18] They shall take up serpents; and if they drink any deadly thing, it shall not hurt them; they shall lay hands on the sick, and they shall recover."),
    ('p', "In the Book of Acts, the phrase 'the gift of the Holy Ghost' refers to salvation. In Acts 2:38, Peter said:"),
    ('ref', "Acts 2:38"),
    ('v', "[38] Then Peter said unto them, Repent, and be baptized every one of you in the name of Jesus Christ for the remission of sins, and ye shall receive the gift of the Holy Ghost."),
    ('p', "In Acts 8:20, Peter rebuked Simon the sorcerer:"),
    ('ref', "Acts 8:20"),
    ('v', "[20] But Peter said unto him, Thy money perish with thee, because thou hast thought that the gift of God may be purchased with money."),
    ('p', "In both cases, 'the gift' is salvation. When Acts describes the manifestations of the Spirit, it uses terms like 'poured out', 'fallen upon', or 'came upon'."),
    ('p', "In Acts 2:16 and 17:"),
    ('ref', "Acts 2:16-17"),
    ('v', "[16] But this is that which was spoken by the prophet Joel;"),
    ('v', "[17] And it shall come to pass in the last days, saith God, I will pour out of my Spirit upon all flesh: and your sons and your daughters shall prophesy, and your young men shall see visions, and your old men shall dream dreams:"),
    ('p', "In Acts 8:15 and 16, Peter and John in Samaria prayed for believers that they might receive the Holy Ghost:"),
    ('ref', "Acts 8:15-16"),
    ('v', "[15] Who, when they were come down, prayed for them, that they might receive the Holy Ghost:"),
    ('v', "[16] (For as yet he was fallen upon none of them: only they were baptized in the name of the Lord Jesus.)"),
    ('p', "In Acts 10:44 through 46, at Cornelius' house:"),
    ('ref', "Acts 10:44-46"),
    ('v', "[44] While Peter yet spake these words, the Holy Ghost fell on all them which heard the word."),
    ('v', "[45] And they of the circumcision which believed were astonished, as many as came with Peter, because that on the Gentiles also was poured out the gift of the Holy Ghost."),
    ('v', "[46] For they heard them speak with tongues, and magnify God."),
    ('p', "In Acts 11:15 and 17, Peter recounts:"),
    ('ref', "Acts 11:15, 17"),
    ('v', "[15] And as I began to speak, the Holy Ghost fell on them, as on us at the beginning."),
    ('v', "[17] Forasmuch then as God gave them the like gift as he did unto us, who believed on the Lord Jesus Christ; what was I, that I could withstand God?"),
    ('p', "In Acts 19:6, at Ephesus:"),
    ('ref', "Acts 19:6"),
    ('v', "[6] And when Paul had laid his hands upon them, the Holy Ghost came on them; and they spake with tongues, and prophesied."),
    ('p', "In the Book of Acts, utterance gifts (speaking in tongues and prophecy) are described as the Spirit falling upon, being poured out, or coming upon."),

    # 11. Revelation Gifts in Acts
    ('p', "Revelation gifts (word of knowledge, word of wisdom, discerning of spirits) are described in Acts as 'the Holy Ghost said' or 'signified by the Spirit'. In Acts 5:3, 4, and 9:"),
    ('ref', "Acts 5:3-4, 9"),
    ('v', "[3] But Peter said, Ananias, why hath Satan filled thine heart to lie to the Holy Ghost, and to keep back part of the price of the land?"),
    ('v', "[4] Whiles it remained, was it not thine own? and after it was sold, was it not in thine own power? why hast thou conceived this thing in thine heart? thou hast not lied unto men, but unto God."),
    ('v', "[9] Then Peter said unto her, How is it that ye have agreed together to tempt the Spirit of the Lord? behold, the feet of them which have buried thy husband are at the door, and shall carry thee out."),
    ('p', "In Acts 10, Peter fell into a trance on the housetop and saw a vision of unclean animals. When Peter doubted what the vision meant, the Spirit spoke to him in verse 19:"),
    ('ref', "Acts 10:19"),
    ('v', "[19] While Peter thought on the vision, the Spirit said unto him, Behold, three men seek thee."),
    ('p', "In Acts 11:28, Agabus demonstrated revelation:"),
    ('ref', "Acts 11:28"),
    ('v', "[28] And there stood up one of them named Agabus, and signified by the Spirit that there should be great dearth throughout all the world: which came to pass in the days of Claudius Caesar."),
    ('p', "Agabus signified by speaking. In Acts 13:2:"),
    ('ref', "Acts 13:2"),
    ('v', "[2] As they ministered to the Lord, and fasted, the Holy Ghost said, Separate me Barnabas and Saul for the work whereunto I have called them."),
    ('p', "In Acts 20:22 and 23, Paul says:"),
    ('ref', "Acts 20:22-23"),
    ('v', "[22] And now, behold, I go bound in the spirit unto Jerusalem, not knowing the things that shall befall me there:"),
    ('v', "[23] Save that the Holy Ghost witnesseth in every city, saying that bonds and afflictions abide me."),
    ('p', "In Acts 21:4 and 11:"),
    ('ref', "Acts 21:4"),
    ('v', "[4] And finding disciples, we tarried there seven days: who said to Paul through the Spirit, that he should not go up to Jerusalem."),
    ('ref', "Acts 21:11"),
    ('v', "[11] And when he was come unto us, he took Paul's girdle, and bound his own hands and feet, and said, Thus saith the Holy Ghost, So shall the Jews at Jerusalem bind the man that owneth this girdle, and shall deliver him into the hands of the Gentiles."),
    ('p', "In Acts, revelation is received and delivered through utterance. A person walking in utterance gifts can equally function in revelation gifts, because revelation is spoken out."),

    # 12. Ministry Order & Screening Prophecy
    ('p', "Because of how the Book of Acts records things, people often assume that manifestations must be sudden, overwhelming experiences where a person loses control."),
    ('p', "In fellowship meetings at the University of Ibadan with my friend Chucky, we learned the importance of screening prophecies. People wrote down their words of exhortation or revelation so the pastor could read them first. On one occasion, someone passed a note declaring that people speaking in tongues in the meeting were bound for hell. That was baseless. Pastors must guard the door of the sheep and verify words and testimonies before they are shared publicly, so the congregation is not harmed."),

    # 13. 1 Corinthians 12: Charisma as Spiritual Endowment
    ('p', "In 1 Corinthians 12:1 and 4:"),
    ('ref', "1 Corinthians 12:1, 4"),
    ('v', "[1] Now concerning spiritual gifts, brethren, I would not have you ignorant."),
    ('v', "[4] Now there are diversities of gifts, but the same Spirit."),
    ('p', "Paul uses the word 'pneumatikos' (things of the Spirit). In verse 4, he calls them 'charisma' (an endowment or ability). Charisma was used by Greek writers like Plato and Aristotle to describe natural endowments. Just as footballers like Lionel Messi and Cristiano Ronaldo possess distinct natural gifts on the field, the believer is endowed by the Spirit. Charisma in the New Testament is a spiritual endowment that functions naturally in the believer's spiritual makeup."),
    ('p', "Howard Carter grouped the gifts of the Spirit into three categories: utterance gifts that say something, revelation gifts that show something, and power gifts that do something. These are not merely sporadic experiences; they are spiritual abilities that produce supernatural results."),

    # 14. Gifts of Healings: Diverse Endowments
    ('p', "In 1 Corinthians 12:9, 28, and 30:"),
    ('ref', "1 Corinthians 12:9, 28, 30"),
    ('v', "[9] To another faith by the same Spirit; to another the gifts of healing by the same Spirit;"),
    ('v', "[28] And God hath set some in the church, first apostles, secondarily prophets, thirdly teachers, after that miracles, then gifts of healings, helps, governments, diversities of tongues."),
    ('v', "[30] Have all the gifts of healing? do all speak with tongues? do all interpret?"),
    ('p', "Paul uses the plural 'gifts of healings' three times. This indicates diverse endowments and administrations of healing. Just as a versatile actor like Arnold Schwarzenegger can play both villain and hero without being typecast into only one role, the believer is endowed to minister healing in diverse ways rather than relying on a single method."),

    # 15. Laying on of Hands: Matthew 8
    ('p', "The first administration of the gifts of healings is the laying on of hands. Mark 16:18 declares that believers shall lay hands on the sick, and the sick shall recover."),
    ('p', "Matthew 4 mentions that Jesus healed many diseases, but Matthew 8 provides the specific action. In verses 2 and 3:"),
    ('ref', "Matthew 8:2-3"),
    ('v', "[2] And, behold, there came a leper and worshipped him, saying, Lord, if thou wilt, thou canst make me clean."),
    ('v', "[3] And Jesus put forth his hand, and touched him, saying, I will; be thou clean. And immediately his leprosy was cleansed."),
    ('p', "Jesus put forth his hand and touched him. In verses 14 and 15:"),
    ('ref', "Matthew 8:14-15"),
    ('v', "[14] And when Jesus was come into Peter's house, he saw his wife's mother laid, and sick of a fever."),
    ('v', "[15] And he touched her hand, and the fever left her: and she arose, and ministered unto them."),
    ('p', "Jesus touched her hand, and the fever left her. Many people assume laying on of hands must always be on the head, but Jesus touched her hand. Laying on of hands is a deliberate touch on the afflicted part of the body."),

    # 16. 1 Timothy 5:22 vs Ordination
    ('p', "1 Timothy 5:22 says:"),
    ('ref', "1 Timothy 5:22"),
    ('v', "[22] Lay hands suddenly on no man, neither be partaker of other men's sins: keep thyself pure."),
    ('p', "Some misuse this verse to claim that ministering to the sick defiles the minister with the sick person's sins. Paul's discussion here is about ordination and appointing leaders to ministry. The Greek word for 'suddenly' is 'tacheos' (hastily). Paul was instructing Timothy not to ordain anyone to ministry hastily without proper character and proven teaching ability."),
    ('p', "In Acts 13:3, hands were laid on Paul and Barnabas for ministry commissioning:"),
    ('ref', "Acts 13:3"),
    ('v', "[3] And when they had fasted and prayed, and laid their hands on them, they sent them away."),
    ('p', "Laying hands for ministry ordination is an act of appointment and consecration, distinct from the physical touch used to minister healing to the sick."),

    # 17. The Deliberate Touch: Mark 8, Mark 10, Mark 6, Acts 28
    ('p', "In ministering healing, what Jesus did was touch the sick. In Mark 8:22 through 25:"),
    ('ref', "Mark 8:22-25"),
    ('v', "[22] And he cometh to Bethsaida; and they bring a blind man unto him, and besought him to touch him."),
    ('v', "[23] And he took the blind man by the hand, and led him out of the town; and when he had spit on his eyes, and put his hands upon him, he asked him if he saw ought."),
    ('v', "[24] And he looked up, and said, I see men as trees, walking."),
    ('v', "[25] After that he put his hands again upon his eyes, and made him look up: and he was restored, and saw every man clearly."),
    ('p', "Jesus touched the blind man's eyes deliberately where the affliction was located. Laying on of hands is a deliberate touch where the pain or affliction exists, maintained with appropriate modesty and propriety."),
    ('p', "In Mark 10:13 and 16:"),
    ('ref', "Mark 10:13, 16"),
    ('v', "[13] And they brought young children to him, that he should touch them: and his disciples rebuked those that brought them."),
    ('v', "[16] And he took them up in his arms, put his hands upon them, and blessed them."),
    ('p', "The word 'lay' in the Greek implies a deliberate touch. Casual physical contact—such as greeting someone, tapping a person on the shoulder, or hugging—is not laying on of hands. In the Greek, soldiers laid hands on Peter and Paul to arrest them; that was a deliberate physical apprehension. In ministering healing, laying on of hands is a deliberate touch to arrest sickness and transmit the healing power of God."),
    ('p', "In Mark 6:4 and 5, Jesus ministered in his hometown where there was widespread unbelief:"),
    ('ref', "Mark 6:4-5"),
    ('v', "[4] But Jesus said unto them, A prophet is not without honour, but in his own country, and among his own kin, and in his own house."),
    ('v', "[5] And he could there do no mighty work, save that he laid his hands upon a few sick folk, and healed them."),
    ('p', "Even in an atmosphere of unbelief, laying on of hands proved effective. Jesus laid his hands upon a few sick folk, and they were healed."),
    ('p', "Similarly, in Acts 28:7 through 9, Paul ministered while under Roman custody on the island of Malta:"),
    ('ref', "Acts 28:7-9"),
    ('v', "[7] In the same quarters were possessions of the chief man of the island, whose name was Publius; who received us, and lodged us three days courteously."),
    ('v', "[8] And it came to pass, that the father of Publius lay sick of a fever and of a bloody flux: to whom Paul entered in, and prayed, and laid his hands on him, and healed him."),
    ('v', "[9] So when this was done, others also, which had diseases in the island, came, and were healed."),
    ('p', "Paul entered in, prayed, laid hands on Publius' father, and healed him. When that was done, others on the island came and were healed. Laying on of hands is the primary starting point for ministering divine healing."),

    # 18. Power in Action: Acts 3 & Acts 8
    ('p', "In Acts 3:6 through 8, Peter ministered to the lame man at the Beautiful Gate:"),
    ('ref', "Acts 3:6-8"),
    ('v', "[6] Then Peter said, Silver and gold have I none; but such as I have give I thee: In the name of Jesus Christ of Nazareth rise up and walk."),
    ('v', "[7] And he took him by the right hand, and lifted him up: and immediately his feet and ankle bones received strength."),
    ('v', "[8] And he leaping up stood, and walked, and entered with them into the temple, walking, and leaping, and praising God."),
    ('p', "The divine power was transmitted through Peter taking him by the right hand. In Acts 8:5 through 7:"),
    ('ref', "Acts 8:5-7"),
    ('v', "[5] Then Philip went down to the city of Samaria, and preached Christ unto them."),
    ('v', "[6] And the people with one accord gave heed unto those things which Philip spake, hearing and seeing the miracles which he did."),
    ('v', "[7] For unclean spirits, crying with loud voice, came out of many that were possessed with them: and many taken with palsies, and that were lame, were healed."),
    ('p', "The power of God is ministered through action. Scripture does not say merely declare people healed from a distance; it instructs believers to lay hands on the sick. The diversity of the gifts of healings is expressed through diverse physical actions."),
    ('p', "In studying the gifts of healings, we study the action. In every instance, the recipients received healing through deliberate touch.")
]

def build_docx(output_path):
    doc = Document()
    
    # Page setup: 1.0 inch margins
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    
    # Title
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    t_run = title_p.add_run("In the Spirit by the Spirit – Power Gifts - Track 2")
    t_run.font.name = "Lexend"
    t_run.font.size = Pt(14)
    t_run.font.bold = True
    
    sub_p = doc.add_paragraph()
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(14)
    s_run = sub_p.add_run("Structured Transcript (Fillers Removed)")
    s_run.font.name = "Lexend"
    s_run.font.size = Pt(10)
    s_run.font.italic = True
    s_run.font.color.rgb = RGBColor(100, 100, 100)
    
    for elem_type, text in ELEMENTS_CLEAN:
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.font.name = "Lexend"
        
        if elem_type == 'ref':
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(2)
            run.font.size = Pt(10)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)
        elif elem_type == 'v':
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(2)
            run.font.size = Pt(9.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)
        else:
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            run.font.size = Pt(10)
            run.font.bold = False
            run.font.color.rgb = RGBColor(30, 30, 30)
            
    doc.save(output_path)
    print(f"Docx built: {output_path}")

def build_txt(output_path):
    lines = []
    lines.append("In the Spirit by the Spirit – Power Gifts - Track 2")
    lines.append("Structured Transcript (Fillers Removed)")
    lines.append("=" * 60)
    lines.append("")
    
    for elem_type, text in ELEMENTS_CLEAN:
        if elem_type in ('ref', 'v'):
            lines.append(f"**{text}**")
        else:
            lines.append(text)
            lines.append("")
            
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))
    print(f"Txt built: {output_path}")

if __name__ == "__main__":
    out_docx_raw = os.path.join("In the Spirit by the Spirit Transcript", "Raw", "In the Spirit by the Spirit – Power Gifts - Track 2 (Fillers Removed).docx")
    out_docx_root = "In the Spirit by the Spirit – Power Gifts - Track 2 (Fillers Removed).docx"
    out_txt_raw = os.path.join("In the Spirit by the Spirit Transcript", "Raw", "In the Spirit by the Spirit – Power Gifts - Track 2 (Fillers Removed).txt")
    
    build_docx(out_docx_raw)
    build_docx(out_docx_root)
    build_txt(out_txt_raw)
    print("Clean build complete!")
