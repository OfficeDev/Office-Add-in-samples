using System;
using System.Reflection;

namespace Office_Add_in_ASPNET_SSO_WebAPI.Areas.HelpPage.ModelDescriptions
{
    public interface IModelDocumentationProvider
    {
        string GetDocumentation(MemberInfo member);

        string GetDocumentation(Type type);
    }
}