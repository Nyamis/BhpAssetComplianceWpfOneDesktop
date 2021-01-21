using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Windows.Input;

namespace BhpAssetComplianceWpfOneDesktop.ViewModels.DataTemplate
{
    public class RibbonButtonDataTemplateViewModel : BindableBase
    {
        private readonly Action _navigationAction;

        private string _myHeader;
        public string MyHeader
        {
            get { return _myHeader; }
            set { SetProperty(ref _myHeader, value); }
        }

        public ICommand PushCommand { get; set; }

        public RibbonButtonDataTemplateViewModel(string header, Action navigationAction)
        {
            _navigationAction = navigationAction;
            PushCommand = new DelegateCommand(navigationAction);
            MyHeader = header;
        }
    }
}
