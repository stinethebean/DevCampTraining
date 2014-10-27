using Microsoft.Office365.OAuth;
using MVCResearchTracker.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace MVCResearchTracker.Controllers
{
    public class ProjectController : Controller
    {
        public async Task<ActionResult> Index(ViewModel submitModel)
        {

            ProjectRepository repo = new ProjectRepository();

            if (submitModel.Project == null)
            {

                ViewModel formModel = repo.GetFromCache("ViewModel") as ViewModel;
                if (formModel == null)
                {
                    formModel = new ViewModel();
                    repo.SaveInCache("ViewModel", formModel);
                }

                try
                {

                    if (formModel.Contacts == null)
                    {
                        formModel.Contacts = await repo.GetMyContacts();
                    }

                    if (formModel.Files == null)
                    {
                        formModel.Files = await repo.GetMyFiles();
                    }

                    if (repo.GetFromCache("UpdateToken") as string == null)
                    {
                        string accessToken = await repo.GetUpdateAccessToken();
                        repo.SaveInCache("UpdateToken", accessToken);
                    }

                }
                catch (RedirectRequiredException x)
                {
                    return Redirect(x.RedirectUri.ToString());
                }

                return View(formModel);
            }
            else
            {
                string projectTitle = submitModel.Project.Title;
                string projectOwner = submitModel.Project.Owner;
                string documentLink = submitModel.Project.DocumentLink;
                string documentName = (repo.GetFromCache("ViewModel") as ViewModel).Files.Where(f => f.Url == documentLink).First().Name;

                bool success = await repo.CreateProject(projectTitle, projectOwner, documentLink, documentName);

                return Redirect("/Project");
            }

        }


    }
}