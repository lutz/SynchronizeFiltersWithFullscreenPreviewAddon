using SwissAcademic.Citavi;
using SwissAcademic.Citavi.Shell;
using SwissAcademic.Collections;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace SynchronizeFiltersWithFullscreenPreview
{
    public class Addon : CitaviAddOn<MainForm>
    {
        #region Fields

        private Dictionary<ReferenceFilterCollection, MainForm> filters;

        #endregion

        #region Constructors

        public Addon()
        {
            filters = new Dictionary<ReferenceFilterCollection, MainForm>();
        }

        #endregion

        #region Methods

        public override void OnHostingFormLoaded(MainForm mainForm)
        {
            if (mainForm.IsPreviewFullScreenForm)
            {
                SetFilterInPreview(mainForm);
            }
            else
            {
                filters.Add(mainForm.ReferenceEditorFilterSet.Filters, mainForm);
                mainForm.FormClosed += MainForm_FormClosed;
                mainForm.ReferenceEditorFilterSet.Filters.CollectionChanged += Filters_CollectionChanged;
            }
            base.OnHostingFormLoaded(mainForm);
        }

        void SynchronizeFilterChangeWithPreview(MainForm mainForm)
        {
            if (mainForm == null)
                return;
            var list = mainForm.GetFilteredReferences();

            foreach (var previewMainForm in GetPreviewMainForms(mainForm.Project))
            {
                previewMainForm.ReferenceEditorFilterSet.Filters.Clear();
                if (list.Count == 0)
                    continue;

                var referenceFilter = new ReferenceFilter(list, "temp", false);
                previewMainForm.ReferenceEditorFilterSet.Filters.Add(referenceFilter);

            }
        }

        void SetFilterInPreview(MainForm previewMainForm)
        {
            var mainForm = GetMainForm(previewMainForm.Project);
            if (mainForm == null)
                return;

            var filteredReferences = mainForm.GetFilteredReferences();
            previewMainForm.ReferenceEditorFilterSet.Filters.Clear();
            if (filteredReferences.Count == 0) return;

            var referenceFilter = new ReferenceFilter(filteredReferences, "temp", false);
            previewMainForm.ReferenceEditorFilterSet.Filters.Add(referenceFilter);
        }

        MainForm GetMainForm(Project project)
        {
            foreach (var openForm in Application.OpenForms)
            {
                if (openForm is MainForm mainForm && !mainForm.IsPreviewFullScreenForm && mainForm.Project.Equals(project))
                    return mainForm;
            }
            return null;
        }

        IEnumerable<MainForm> GetPreviewMainForms(Project project)
        {
            foreach (var openForm in Application.OpenForms)
            {
                if (openForm is MainForm mainForm && mainForm.IsPreviewFullScreenForm && mainForm.Project.Equals(project))
                    yield return mainForm;
            }
        }

        #endregion

        #region EventHandlers

        void Filters_CollectionChanged(object sender, CollectionChangedEventArgs<ReferenceFilter> e)
        {
            if (sender is ReferenceFilterCollection key)
                SynchronizeFilterChangeWithPreview(filters[key]);
        }

        void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (sender is MainForm mainForm)
            {
                mainForm.FormClosed -= MainForm_FormClosed;
                mainForm.ReferenceEditorFilterSet.Filters.CollectionChanged -= Filters_CollectionChanged;
                filters.RemoveAll(pair => pair.Value.Equals(mainForm));
            }
        }

        #endregion
    }
}