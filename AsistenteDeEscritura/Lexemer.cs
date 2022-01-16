using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AsistenteDeEscritura
{
    class Lexemer
    {
        public enum Region
        {
            R1,
            R2
        }
        struct Step1Info
        {
            public Step1Info(Regex i_regex, string i_replace, Region i_region)
            {
                regex = i_regex;
                replace = i_replace;
                region = i_region;
            }

            public Regex regex;
            public string replace;
            public Region region;
        }
        Regex m_rv_re = new Regex("(.[^aeiouáéíóúü]+[aeiouáéíóúü]|[aeiouáéíóúü]{2,}[^aeiouáéíóúü]|[^aeiouáéíóúü][aeiouáéíóúü].)(.*)");
        //Regex m_r1r2_re = new Regex("^.*?[aeiouáéíóúü][^aeiouáéíóúü](.*)");
        Regex m_r1r2_re = new Regex("^.*?[aeiouáéíóúü][^aeiouáéíóúü](.*)");
        Regex m_step0_re = new Regex("((i[eé]ndo|[aá]ndo|[aáeéií]r|u?yendo)(sel[ao]s?|l[aeo]s?|nos|se|me))$");
        Step1Info[] m_step1_re = new Step1Info[]
        {
            new Step1Info(new Regex("(anzas?|ic[oa]s?|ismos?|[ai]bles?|istas?|os[oa]s?|[ai]mientos?)$"), "", Region.R2 ),
            new Step1Info(new Regex("((ic)?(adora?|ación|ador[ae]s|aciones|antes?|ancias?))$"), "", Region.R2 ),
            new Step1Info(new Regex("(logías?)$"), "log", Region.R2 ),
            new Step1Info(new Regex("(ución|uciones)$"), "u", Region.R2 ),
            new Step1Info(new Regex("(encias?)$"), "ente", Region.R2 ),
            new Step1Info(new Regex("((os|ic|ad|(at)?iv)amente)$"), "", Region.R2 ),
            new Step1Info(new Regex("(amente)$"), "", Region.R1 ),
            new Step1Info(new Regex("((ante|[ai]ble)?mente)$"), "", Region.R2),
            new Step1Info(new Regex("((abil|ic|iv)?idad(es)?)$"), "", Region.R2),
            new Step1Info(new Regex("((at)?iv[ao]s?)$"), "", Region.R2)
        };

        Regex m_step2a_re = new Regex("(y[ae]n?|yeron|yendo|y[oó]|y[ae]s|yais|yamos)$");
        Regex m_step2b_re_1 = new Regex("(en|es|éis|emos)$");
        Regex[] m_step2b_re_2 = new Regex[]
        {
            new Regex("(([aei]ría|ié(ra|se))mos)$"),
            new Regex("(([aei]re|á[br]a|áse)mos)$"),
            new Regex("([aei]ría[ns]|[aei]réis|ie((ra|se)[ns]|ron|ndo)|a[br]ais|aseis|íamos)$"),
            new Regex("([aei](rá[ns]|ría)|a[bdr]as|id[ao]s|íais|([ai]m|ad)os|ie(se|ra)|[ai]ste|aban|ar[ao]n|ase[ns]|ando)$"),
            new Regex("([aei]r[áé]|a[bdr]a|[ai]d[ao]|ía[ns]|áis|ase)$"),
            new Regex("(í[as]|[aei]d|a[ns]|ió|[aei]r)$")
        };
        Regex m_step3_re_1 = new Regex("(os|a|o|á|í|ó)$");
        Regex m_step3_re_2 = new Regex("(u?é|u?e)$");

        struct Regions
        {
            public string r1;
            public string r2;
            public string rv;
            public string word;

            public string Get(Region i_region)
            {
                switch(i_region)
                {
                    case Region.R1:
                        return r1;
                    case Region.R2:
                        return r2;
                }
                return word;
            }
        }

        Regions getRegions(string word)
        {
            //R1 is the region after the first non-vowel following a vowel,
            //or is the null region at the end of the word if there is no
            //such non-vowel.
            Match r1Match = m_r1r2_re.Match(word);
            string r1 = (r1Match != null && r1Match.Success) ? r1Match.Groups[1].Value : "";

            //R2 is the region after the first non-vowel following a vowel
            //in R1, or is the null region at the end of the word if there
            //is no such non-vowel.
            Match r2Match = m_r1r2_re.Match(r1);
            string r2 = (r2Match != null && r2Match.Success) ? r2Match.Groups[1].Value : "";

            //If the second letter is a consonant, RV is the region after
            //the next following vowel, or if the first two letters are
            //vowels, RV is the region after the next consonant, and
            //otherwise (consonant-vowel case) RV is the region after the
            //third letter. But RV is the end of the word if these positions
            //cannot be found.
            Match rvMatch = m_rv_re.Match(word);
            string rv = (rvMatch != null && rvMatch.Success )? rvMatch.Groups[2].Value : "";

            Regions regions = new Regions();
            regions.r1 = r1;
            regions.r2 = r2;
            regions.rv = rv;
            regions.word = word;
            return regions;
        }

        string removeAccents(string i_word)
        {
            return i_word.Replace("á", "a").Replace("é", "e").Replace("í", "i").Replace("ó", "o").Replace("u", "u");
        }

        string step0(Regions i_regions)
        {
            //Search for the longest among the following suffixes 
            //me, se, sela, selo, selas, selos, la, le, lo, las, les, los, nos
            //and delete it, if comes after one of:
            //iéndo, ándo, ár, ér, ír, ando, iendo, ar, er, ir, yendo(following u)
            Match match = m_step0_re.Match(i_regions.rv);
            if(match != null && match.Success)
            {
                GroupCollection g = match.Groups;
                string w = i_regions.word.Substring(0, i_regions.word.Length - g[1].Length);
                //In the case of (yendo following u), yendo must lie in RV,
                //but the preceding u can be outside it.
                if (g[2].Value == "yendo" && w.Substring(w.Length - 1) == "u")
                {
                    return i_regions.word;
                }
                //In the case of (iéndo   ándo   ár   ér   ír), deletion is
                //followed by removing the acute accent (for example,
                //haciéndola -> haciendo).
                return w + removeAccents(g[2].Value);
            }
            else
            {
                return i_regions.word;
            }

        }

        string step2b(Regions r)
        {
            //Search for the longest among the following suffixes in RV,
            //and perform the action indicated.

            //iera, ad, ed, id, ase, iese, aste, iste, an, aban, ían, aran,
            //ieran, asen, iesen, aron, ieron, ado, ido, ando, iendo, ió, ar,
            //er, ir, as, idas, ías, aras, ieras, ases, ieses,
            //áis, abais, íais, arais, ierais,   aseis, ieseis, asteis,
            //isteis, ados, idos, amos, ábamos, íamos, imos, áramos, iéramos,
            //iésemos, ásemos
            //delete
            for (var i = 0; i < m_step2b_re_2.Length; i++)
            {
                Match m = m_step2b_re_2[i].Match(r.rv);
                if (m != null && m.Success)
                {
                    return m_step2b_re_2[i].Replace(r.word, "");
                }
            }

            //en, es, éis, emos
            //delete, and if preceded by gu delete the u (the gu need not be
            // in RV)
            Match match = m_step2b_re_1.Match(r.rv);
            if (match != null && match.Success)
            {
                var w = r.word.Substring(0, r.word.Length - match.Groups[1].Length);
                if (w.Substring(w.Length -2 ) == "gu")
                {
                    return w.Substring(0, w.Length - 1);
                }
                return w;
            }

            return r.word;
        }

        string step2a(Regions r)
        {
            //Search for the longest among the following suffixes in RV,
            //and if found, delete if preceded by u.
            //ya, ye, yan, yen, yeron, yendo, yo, yó, yas, yes, yais, yamos
            Match match = m_step2a_re.Match(r.rv);

            if (match != null && match.Success)
            {
                GroupCollection g = match.Groups;
                string w = r.word.Substring(0, r.word.Length - g[1].Length);

                //Note that the preceding u need not be in RV
                if (w.Substring(w.Length - 1) == "u")
                {
                    return w;
                }
            }

            //Do Step 2b if step 2a was done, but failed to remove a suffix.
            return step2b(r);
        }

        string step1(Regions r)
        {
            for(int i = 0; i < m_step1_re.Length; i++)
            {
                Step1Info rule = m_step1_re[i];
                Match match = rule.regex.Match(r.Get(rule.region));
                if(match != null && match.Success)
                {
                    GroupCollection g = match.Groups;
                    string w = r.word.Substring(0, r.word.Length - g[1].Length);
                    return w + rule.replace;
                }
            }

            return step2a(r);
        }

        string step3(Regions r)
        {
            //Search for the longest among the following suffixes in RV, and
            //perform the action indicated.

            //os, a, o, á, í, ó
            //delete if in RV 
            var match = m_step3_re_1.Match(r.rv);
            if (match != null && match.Success)
            {
                return m_step3_re_1.Replace(r.word, "");
            }

            //e, é
            //delete if in RV, and if preceded by gu with the u in RV delete
            //the u
            match = m_step3_re_2.Match(r.rv);
            if (match != null && match.Success)
            {
                if (match.Groups[1].Value.Substring(0, 1) == "u" && r.word.Substring(r.word.Length -3, 2) == "gu")
                {
                    return r.word.Substring(0, r.word.Length - 2);
                }
                return r.word.Substring(0, r.word.Length - 1);
            }

            return r.word;
        }

        public string ComputeLexema(string i_word)
        {
            i_word = i_word.ToLower().Trim();
            Regions r0 = getRegions(i_word);
            string word0 = step0(r0);
            Regions r1 = getRegions(word0);
            string word1 = step1(r1);
            Regions r3 = getRegions(word1);
            string word3 = step3(r3);
            return removeAccents(word3);
        }
    }
}
