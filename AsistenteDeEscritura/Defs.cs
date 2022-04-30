using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AsistenteDeEscritura
{
    public interface IFlagReason
    {

    }

    class FlagReasonRepetition: IFlagReason
    {

    }

    class FlagReasonRepetitionClose: FlagReasonRepetition
    {

    }

    class FlagReasonRepetitionDistant : FlagReasonRepetition
    {

    }

    class FlagReasonRime : IFlagReason
    {

    }

    class FlagReasonRimeConsonant : FlagReasonRime
    {

    }

    class FlagReasonRimeAsonant : FlagReasonRime
    {

    }

    class FlagReasonAdverbioMente : IFlagReason
    {

    }

    class FlagReasonGerundio : IFlagReason
    {

    }

    class FlagReasonGuion : IFlagReason
    {

    }

    class FlagReasonDiciente : IFlagReason
    {

    }

    class FlagReasonAdjetivo : IFlagReason
    {

    }

    class FlagReasonCacofonia : IFlagReason
    {

    }

    class FlagReasonFraseLarga : IFlagReason
    {
        public FlagReasonFraseLarga(int i_size)
        {

        }
    }
}
