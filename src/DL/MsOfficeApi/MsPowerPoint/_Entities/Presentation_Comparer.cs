using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JBC.ExploreTheWorld.DL.MsOfficeApi.MsPowerPoint
{
    public class Presentation_Comparer
    {
        public Presentation Source_Presentation { get; private set; }
        public Presentation Target_Presentation { get; private set; }

        public Presentation_Comparer(Presentation source_Presentation, Presentation target_Presentation)
        {
            Source_Presentation = source_Presentation;
            Target_Presentation = target_Presentation;

            Initialize();
        }

        public void Initialize()
        {
        }

        List<int> _SlidesWithUpdatedNotes = new List<int>();

        public void Compare()
        {
            var sourceSlides = Source_Presentation.Slides ?? new List<Slide>();
            var targetSlides = Target_Presentation.Slides ?? new List<Slide>();

            foreach (var sourceSlide in sourceSlides)
            {
                var targetSlide = targetSlides.Where(o => o.SlideNumber.Equals(sourceSlide.SlideNumber)).FirstOrDefault();

                if (targetSlide != null)
                {
                    if (SlideHasNarration(targetSlide) == false)
                    {
                        _SlidesWithUpdatedNotes.Add(targetSlide.SlideNumber.GetValueOrDefault());
                        continue;
                    }

                    var sourceSlideNotes = sourceSlide.GetNotes();
                    var targetSlideNotes = targetSlide.GetNotes();

                    if (sourceSlideNotes != targetSlideNotes)
                    {
                        _SlidesWithUpdatedNotes.Add(targetSlide.SlideNumber.GetValueOrDefault());
                    }
                }
            }
        }

        public bool SlideHasNarration(Slide slide)
        {
            if (slide.Shapes == null) return false;

            var shape = slide.Shapes.Where(o => o.IsNarrationAsBoolean() == true).FirstOrDefault();

            if (shape != null) return true;

            return false;
        }

        public List<int> GetSlidesWithUpdatedNotes()
        {
            return _SlidesWithUpdatedNotes;
        }
    }
}
