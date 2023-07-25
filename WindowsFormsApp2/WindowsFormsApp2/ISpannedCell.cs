using System.Windows.Forms;

namespace WindowsFormsApp2
{
    interface ISpannedCell
    {
        int ColumnSpan { get; }
        int RowSpan { get; }
        DataGridViewCell OwnerCell { get; }
    }
}