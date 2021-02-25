using BhpAssetComplianceWpfOneDesktop.ViewModels.DataTemplate;
using Prism.Mvvm;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels
{
    public abstract class BasePosterViewModel : BindableBase
    {
        protected abstract string MyPosterName { get; set; }
        protected abstract string MyPosterIcon { get; set; }


        private PosterHeaderDataTemplateViewModel _myPosterHeader;
        public PosterHeaderDataTemplateViewModel MyPosterHeader
        {
            get { return _myPosterHeader; }
            set { SetProperty(ref _myPosterHeader, value); }
        }

        protected BasePosterViewModel()
        {
            MyPosterHeader = new PosterHeaderDataTemplateViewModel(MyPosterName,MyPosterIcon);
        }
    }
}
