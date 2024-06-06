                    def masterFun(a):
                        try:
                            penalty_code = a
                            if penalty_code:
                                age_part, handicap_part = penalty_code.split("(-")
                                av_age = age_part.split()[1]
                                handicap = handicap_part.split(")")[0]
                                return(f'AV AGE: {av_age} <br> HANDICAP: {handicap}')
                            else:
                                return(" ")
                        except IndexError:
                            pass

                    data.append(
                        [str(fl["Place"][j]).split(sep=".")[0], str(fl["Bow"][j]).split(sep=".")[0], f'<img src="flags/{flag_list[fl["CrewAbbrev"][j]]}" style="max-width: 6mm">',
                         fl["Crew"][j],
                         fl["Stroke"][j].replace("/", "<br>"), fl["AdjTime"][j], fl["Delta"][j], " ", " ", en])
                    dataQ.append(
                        [str(fl["Place"][j]).split(sep=".")[0], f"({str(fl["Rank"][j]).split(sep='.')[0]})", str(fl["Bow"][j]).split(sep=".")[0], f'<img src="flags/{flag_list[fl["CrewAbbrev"][j]]}" style="max-width: 6mm">',
                         fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), fl["AdjTime"][j], fl["Delta"][j], " ", " ", en])

                    dataMaster.append(
                        [str(fl["Place"][j]).split(sep=".")[0], str(fl["Bow"][j]).split(sep=".")[0], f'<img src="flags/{flag_list[fl["CrewAbbrev"][j]]}" style="max-width: 6mm">',
                         fl["Crew"][j],
                         fl["Stroke"][j].replace("/", "<br>"), fl["RawTime"][j], fl["AdjTime"][j], fl["Delta"][j], " ", " ",
                         fl["Qual"][j], masterFun(str(fl["PenaltyCode"][j])),  en])

                    dataMasterQ.append(
                        [str(fl["Place"][j]).split(sep=".")[0], f"({str(fl["Rank"][j]).split(sep='.')[0]})",
                         str(fl["Bow"][j]).split(sep=".")[0],
                         f'<img src="flags/{flag_list[fl["CrewAbbrev"][j]]}" style="max-width: 6mm">',
                         fl["Crew"][j], fl["Stroke"][j].replace("/", "<br>"), fl["RawTime"][j], fl["AdjTime"][j], fl["Delta"][j], " ",
                         " ", masterFun(str(fl["PenaltyCode"][j])), en])

                    start_data.append(
                        [str(fl["Bow"][j]).split(sep=".")[0], f'<img src="flags/{flag_list[fl["CrewAbbrev"][j]]}" style="max-width: 6mm">', fl["Crew"][j], fl["Stroke"][j].replace("/", ", "), masterFun(str(fl["PenaltyCode"][j])), en])
