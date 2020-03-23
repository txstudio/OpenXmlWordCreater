using System.Collections.Generic;

namespace RegistrationFormCreater
{
    public sealed class Form
    {
        public string TeamName { get; set; }
        public string Level { get; set; }

        public string LeaderName { get; set; }
        public string LeaderPhone { get; set; }

        public string CoachName { get; set; }
        public string CoachPhone { get; set; }

        public IEnumerable<PlayerItem> Players { get; set; }
    }

    public sealed class PlayerItem
    {
        public string Number { get; set; }
        public string Name { get; set; }
        public bool IsCaption { get; set; }
        public bool IsGoalKeeper { get; set; }
        public string Note { get; set; }

        public string CaptionMark
        {
            get
            {
                if (this.IsCaption == true)
                    return "V";

                return string.Empty;
            }
        }

        public string GoalKeeperMark
        {
            get
            {
                if (this.IsGoalKeeper == true)
                    return "V";

                return string.Empty;
            }
        }
    }

}
