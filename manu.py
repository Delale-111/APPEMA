import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="APP MANU", page_icon="📄", layout="wide")

# Chemin vers le fichier Excel
fichier_path = 'fichier.xlsx'

def load_data():
    return pd.read_excel(fichier_path)

def update_excel(chapitre, termine):
    book = openpyxl.load_workbook(fichier_path)
    sheet = book.active
    
    # Trouver la ligne correspondante au chapitre
    for row in range(2, sheet.max_row + 1):
        if sheet[f'A{row}'].value == chapitre:
            sheet[f'B{row}'] = termine
            break
    
    book.save(fichier_path)
    book.close()
    return load_data()  # Retourne les données mises à jour

def update_progression(data):
    # Mise à jour de la progression dans st.session_state
    st.session_state.progression = sum(data["Terminé"] * data["Pourcentage évolution"])

# Initialisation de la progression avec les données actuelles du fichier Excel
if 'progression' not in st.session_state:
    data_initiale = load_data()
    update_progression(data_initiale)

def main():
    st.title("Explorateur de Chapitres")
    data = load_data()

    col1, col2 = st.columns([3, 1])

    with col1:
        chapitre_choice = st.selectbox("Choisissez un chapitre", data["CHAPITRES"])
        chapitre_info = data[data["CHAPITRES"] == chapitre_choice].iloc[0]
        
        # Affichage simulé du contenu du chapitre
        st.write(f"Contenu du Chapitre {chapitre_choice}")
        if chapitre_choice == 1:
            st.write(":red[01] LA TROISIEME ANNEE du règne de Joakim, roi de Juda, Nabucodonosor, roi de Babylone, arriva devant Jérusalem et l’assiégea.")
            st.write(":red[02] Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write(":red[03] Le roi ordonna à Ashpénaz, chef de ses eunuques, de faire venir quelques jeunes Israélites de race royale ou de famille noble.")
            st.write(":red[04] Ils devaient être sans défaut corporel, de belle figure, exercés à la sagesse, instruits et intelligents, pleins de vigueur, pour se tenir à la cour du roi et apprendre l’écriture et la langue des Chaldéens.")
            st.write(":red[05] Le roi leur assignait pour chaque jour une portion des mets royaux et du vin de sa table. Ils devaient être formés pendant trois ans, et ensuite ils entreraient au service du roi.")
            st.write(":red[06] Parmi eux se trouvaient Daniel, Ananias, Misaël et Azarias, qui étaient de la tribu de Juda.")
            st.write(":red[07] Le chef des eunuques leur imposa des noms : à Daniel celui de Beltassar, à Ananias celui de Sidrac, à Misaël celui de Misac, et à Azarias celui d’Abdénago.")
            st.write(":red[08] Daniel eut à cœur de ne pas se souiller avec les mets du roi et le vin de sa table, il supplia le chef des eunuques de lui épargner cette souillure.")
            st.write(":red[09] Dieu permit à Daniel de trouver auprès de celui-ci faveur et bienveillance.")
            st.write(":red[10] Mais il répondit à Daniel : « J’ai peur de mon Seigneur le roi, qui a fixé votre nourriture et votre boisson ; s’il vous voit le visage plus défait qu’aux jeunes gens de votre âge, c’est moi qui, à cause de vous, risquerai ma tête devant le roi. »")
            st.write(":red[11] Or, le chef des eunuques avait confié Daniel, Ananias, Azarias et Misaël à un intendant. Daniel lui dit :")
            st.write(":red[12] « Fais donc pendant dix jours un essai avec tes serviteurs : qu’on nous donne des légumes à manger et de l’eau à boire.")
            st.write(":red[13] Tu pourras comparer notre mine avec celle des jeunes gens qui mangent les mets du roi, et tu agiras avec tes serviteurs suivant ce que tu auras constaté. »")
            st.write(":red[14] L’intendant consentit à leur demande, et les mit à l’essai pendant dix jours.")
            st.write(":red[15] Au bout de dix jours, ils avaient plus belle mine et meilleure santé que tous les jeunes gens qui mangeaient des mets du roi.")
            st.write(":red[16] L’intendant supprima définitivement leurs mets et leur ration de vin, et leur fit donner des légumes.")
            st.write(":red[17] À ces quatre jeunes gens, Dieu accorda science et habileté en matière d’écriture et de sagesse. Daniel, en outre, savait interpréter les visions et les songes.")
            st.write(":red[18] Au terme fixé par le roi Nabucodonosor pour qu’on lui amenât tous les jeunes gens, le chef des eunuques les conduisit devant lui.")
            st.write(":red[19] Le roi s’entretint avec eux, et pas un seul n’était comparable à Daniel, Ananias, Misaël et Azarias. Ils entrèrent donc au service du roi.")
            st.write(":red[20] Sur toutes les questions demandant sagesse et intelligence que le roi leur posait, il les trouvait dix fois supérieurs à tous les magiciens et mages de tout son royaume.")
            st.write(":red[21] Et Daniel vécut jusqu’à la première année du roi Cyrus.")

        elif chapitre_choice == 2:
            st.write(":red[01] La deuxième année de son règne, Nabucodonosor eut des songes, et le sommeil quitta son esprit troublé.")
            st.write(":red[02] Le roi fit appeler les magiciens, les mages, les enchanteurs et les devins, pour qu’ils interprètent les songes du roi. Ils arrivèrent et se tinrent en présence du roi.")
            st.write(":red[03] Le roi leur dit : « J’ai eu un songe, et mon esprit est troublé par le désir de le comprendre. »")
            st.write(":red[04] Les devins dirent au roi en araméen : « Ô roi, puisses-tu vivre à jamais ! Raconte le songe à tes serviteurs, et nous en donnerons l’interprétation. »")
            st.write(":red[05] Le roi répondit aux devins : « Je n’ai qu’une parole ! Faites-moi connaître le songe et son interprétation, sinon vous serez mis en pièces et vos maisons ne seront plus que décombres.")
            st.write(":red[06] Par contre, si vous me faites connaître le songe et son interprétation, vous recevrez de moi des cadeaux, des récompenses et de grands honneurs. Faites-moi donc connaître le songe et son interprétation. »")
            st.write(":red[07] Pour la deuxième fois, ils répondirent : « Que le roi dise le songe à ses serviteurs, et nous en ferons connaître l’interprétation. »")
            st.write(":red[08] Mais le roi répondit : « Bien entendu, vous cherchez à gagner du temps, maintenant que j’ai donné ma parole !")
            st.write(":red[09] Si vous ne me faites pas connaître le songe, il n’y aura pour vous qu’une seule sentence. Vous êtes complices et vous me tenez des discours mensongers et retors pour tromper le temps. Racontez-moi le songe, et je saurai que vous m’en ferez connaître l’interprétation. »")
            st.write(":red[10] Les devins répondirent au roi : « Personne au monde ne peut faire connaître ce que demande le roi. D’ailleurs, aucun roi, si grand et si puissant soit-il, n’a encore demandé une chose pareille à un magicien, un mage ou un devin.")
            st.write(":red[11] Ce que le roi demande est si difficile que seuls les dieux, dont la demeure n’est pas parmi les hommes, pourraient le faire connaître au roi. »")
            st.write(":red[12] Alors, le roi laissa exploser une terrible colère, et il ordonna de faire exécuter tous les sages de Babylone.")
            st.write(":red[13] La condamnation à mort des sages fut promulguée, et l’on fit chercher Daniel et ses compagnons pour les faire mourir.")
            st.write(":red[14] Mais Daniel, en des paroles sages et prudentes, s’adressa à Aryok, chef des gardes du roi, qui s’apprêtait à faire mourir les sages de Babylone.")
            st.write(":red[15] Il parla ainsi à Aryok, l’officier du roi : « Pourquoi la sentence du roi est-elle si dure ? » Et Aryok l’expliqua à Daniel.")
            st.write(":red[16] Daniel alla demander au roi de lui accorder un délai pour faire connaître au roi l’interprétation du songe.")
            st.write(":red[17] Puis, Daniel retourna chez lui et mit au courant de l’affaire Ananias, Misaël et Azarias, ses compagnons.")
            st.write(":red[18] Il leur demanda d’implorer la miséricorde du Dieu du ciel à propos de ce mystère, pour qu’on n’exécute pas Daniel et ses compagnons avec les autres sages de Babylone.")
            st.write(":red[19] Alors, dans une vision nocturne, le mystère fut révélé à Daniel. Et Daniel bénit le Dieu du ciel.")
            st.write(":red[20] Daniel prit la parole et dit : « Béni soit le nom de Dieu depuis toujours et à jamais. À lui la sagesse et la force !")
            st.write(":red[21] Lui qui fait changer les âges et les temps, il renverse des rois, il en établit d’autres ; aux sages il donne la sagesse, et l’intelligence à ceux qui savent discerner.")
            st.write(":red[22] Lui qui révèle profondeurs et secrets, il connaît ce qui est dans les ténèbres, et la lumière demeure auprès de lui.")
            st.write(":red[23] À toi, Dieu de mes pères, mon action de grâce et ma louange, car tu m’as donné la sagesse et la force, et maintenant tu m’as fait connaître ce que nous t’avons demandé, puisque tu nous as fait connaître ce qui concerne le roi. »")
            st.write(":red[24] Après quoi, Daniel alla chez Aryok, que le roi avait chargé d’exécuter les sages de Babylone. Il entra et lui parla ainsi : « N’exécute pas les sages de Babylone ! Conduis-moi devant le roi. Je ferai connaître au roi l’interprétation du songe. »")
            st.write(":red[25] En toute hâte, Aryok conduisit Daniel devant le roi et lui parla ainsi : « Parmi les déportés de Juda, j’ai trouvé un homme qui donnera l’interprétation au roi. »")
            st.write(":red[26] Prenant la parole, le roi dit à Daniel, surnommé Beltassar : « Peux-tu me faire connaître ce que j’ai vu en songe et son interprétation ? »")
            st.write(":red[27] En présence du roi, Daniel répondit : « Le mystère sur lequel le roi s’interroge, des sages, des mages, des magiciens ou des astrologues ne peuvent le faire connaître au roi.")
            st.write(":red[28] Mais, dans les cieux, il y a un Dieu qui révèle les mystères et fait connaître au roi Nabucodonosor ce qui arrivera à la fin des jours. Ton songe et les visions de ton esprit sur ton lit, les voici.")
            st.write(":red[29] Ô roi, sur ton lit, des pensées ont surgi à ton esprit au sujet de ce qui arrivera par la suite. Celui qui révèle les mystères t’a fait connaître ce qui arrivera.")
            st.write(":red[30] Quant à moi, ce n’est pas à cause d’une sagesse qui, en moi, serait supérieure à celle de tout être vivant, que le mystère m’a été révélé ; mais c’est afin que l’on fasse connaître au roi l’interprétation, et que tu connaisses les pensées de ton cœur.")
            st.write(":red[31] Ô roi, voici ta vision : une énorme statue se dressait devant toi, une grande statue, extrêmement brillante et d’un aspect terrifiant.")
            st.write(":red[32] Elle avait la tête en or fin ; la poitrine et les bras, en argent ; le ventre et les cuisses, en bronze ;")
            st.write(":red[33] ses jambes étaient en fer, et ses pieds, en partie de fer, en partie d’argile.")
            st.write(":red[34] Tu étais en train de regarder : soudain une pierre se détacha d’une montagne, sans qu’on y ait touché ; elle vint frapper les pieds de fer et d’argile de la statue et les pulvérisa.")
            st.write(":red[35] Alors furent pulvérisés tout ensemble le fer et l’argile, le bronze, l’argent et l’or ; ils devinrent comme la paille qui s’envole en été, au moment du battage : ils furent emportés par le vent sans laisser de traces. Quant à la pierre qui avait frappé la statue, elle devint un énorme rocher qui remplit toute la terre.")
            st.write(":red[36] Voici le songe ; et maintenant, en présence du roi, nous allons en donner l’interprétation.")
            st.write(":red[37] C’est à toi, le roi des rois, que le Dieu du ciel a donné royauté, puissance, force et gloire.")
            st.write(":red[38] C’est à toi qu’il a remis les enfants des hommes, les bêtes des champs et les oiseaux du ciel, quelle que soit leur demeure ; c’est toi qu’il a rendu maître de toute chose : la tête d’or, c’est toi.")
            st.write(":red[39] Après toi s’élèvera un autre royaume inférieur au tien, ensuite un troisième royaume, un royaume de bronze qui dominera la terre entière.")
            st.write(":red[40] Il y aura encore un quatrième royaume, dur comme le fer. De même que le fer brise et écrase tout, de même, il pulvérisera et brisera tous les royaumes.")
            st.write(":red[41] Tu as vu les pieds qui étaient en partie d’argile et en partie de fer : en effet, ce royaume sera divisé ; il aura en lui la force du fer, comme tu as vu du fer mêlé à l’argile.")
            st.write(":red[42] Ces pieds en partie de fer et en partie d’argile signifient que le royaume sera en partie fort et en partie faible.")
            st.write(":red[43] Tu as vu le fer associé à l’argile parce que les royaumes s’uniront par des mariages ; mais ils ne tiendront pas ensemble, de même que le fer n’adhère pas à l’argile.")
            st.write(":red[44] Or, au temps de ces rois, le Dieu du ciel suscitera un royaume qui ne sera jamais détruit, et dont la royauté ne passera pas à un autre peuple. Ce dernier royaume pulvérisera et anéantira tous les autres, mais lui-même subsistera à jamais.")
            st.write(":red[45] C’est ainsi que tu as vu une pierre se détacher de la montagne sans qu’on y ait touché, et pulvériser le fer, le bronze, l’argile, l’argent et l’or. Le grand Dieu a fait connaître au roi ce qui doit ensuite advenir. Le songe disait vrai, l’interprétation est digne de foi. »")
            st.write(":red[46] Alors, le roi Nabucodonosor tomba face contre terre. Se prosternant devant Daniel, il ordonna qu’on lui présente une offrande de céréales et un sacrifice d’agréable odeur.")
            st.write(":red[47] Le roi prit la parole et dit à Daniel : « En vérité, votre Dieu est le Dieu des dieux, le Seigneur des rois, celui qui révèle les mystères, puisque tu as su nous révéler ce mystère. »")
            st.write(":red[48] Puis le roi conféra un rang élevé à Daniel et lui offrit de riches et nombreux cadeaux. Il lui donna autorité sur toute la province de Babylone et en fit le préfet suprême de tous les sages de Babylone.")
            st.write(":red[49] Daniel demanda au roi de confier l’administration de la province de Babylone à Sidrac, Misac et Abdénago. Quant à Daniel, il était à la cour du roi.")

        elif chapitre_choice == 3:
            st.write("02 Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write("et ainsi de suite...")

        elif chapitre_choice == 4:
            st.write("02 Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write("et ainsi de suite...")

        elif chapitre_choice == 5:
            st.write("02 Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write("et ainsi de suite...")

        elif chapitre_choice == 6:
            st.write("02 Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write("et ainsi de suite...")

        elif chapitre_choice == 7:
            st.write("02 Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write("et ainsi de suite...")

        elif chapitre_choice == 8:
            st.write("02 Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write("et ainsi de suite...")

        elif chapitre_choice == 9:
            st.write("02 Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write("et ainsi de suite...")

        elif chapitre_choice == 10:
            st.write("02 Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write("et ainsi de suite...")

        elif chapitre_choice == 11:
            st.write("02 Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write("et ainsi de suite...")

        elif chapitre_choice == 12:
            st.write("02 Le Seigneur livra entre ses mains Joakim, roi de Juda, ainsi qu’une partie des objets de la Maison de Dieu. Il les emporta au pays de Babylone, et les déposa dans le trésor de ses dieux.")
            st.write("et ainsi de suite...")
        
        termine = st.checkbox("Terminé", chapitre_info["Terminé"] == 1)


        if st.button("Mettre à jour"):
            updated_data = update_excel(chapitre_choice, int(termine))
            update_progression(updated_data)
            st.success("Mise à jour effectuée.")

    with col2:
        st.write("Votre progression :")
        # Affichage de la jauge de progression sans division par 100
        st.progress(st.session_state.progression)
        st.write(f"{st.session_state.progression:.2f}%")

main()
