using SpResearchTracker.Filters;
using SpResearchTracker.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace SpResearchTracker.Controllers
{
    [Authorize]
    [OAuthExceptionFilter]
    public class ConfigurationsController : ApiController
    {
        //This interface is used to support dependency injection
        private IResearchRepository _repository;

        public ConfigurationsController(IResearchRepository repository)
        {
            _repository = repository;
        }

        /// <summary>
        /// Gets information about the current list configuration in SharePoint.
        /// If the appropriate lists do not exist, then they are created.
        /// Client applications can call this controller at startup to ensure
        /// SharePoint is properly configured.
        /// </summary>
        /// <returns>A collection of list names</returns>
        public async Task<IHttpActionResult> Get()
        {
            //Get access token to SharePoint
            string accessToken = ((Repository)_repository).GetAccessToken();
            if (accessToken == null)
            {
                throw new UnauthorizedAccessException();
            }

            //Check to see if the lists exist
            bool projectsListExists = await _repository.ListExists(accessToken, ((Repository)_repository).ProjectsListName);

            //Create the "Projects" list, if necessary
            if (!projectsListExists)
            {
                bool projectsListCreated = await _repository.CreateList(accessToken, ((Repository)_repository).ProjectsListName, "100");
            }

            //Return the names of the lists as validation that the configuration is correct
            List<ConfigurationInfo> configurations = await _repository.GetConfigurations(accessToken);
            return Ok<List<ConfigurationInfo>>(configurations); 

        }
    }
}
