"""
This module processes web requests from WIKI to DOC page,
it reads a wiki page and prints this information into a word
document (docx/docm).
"""

import re
import urllib
from genshi.builder import tag
from trac.core import Component, implements
from trac.web import IRequestHandler
from trac.web.chrome import INavigationContributor, ITemplateProvider, add_stylesheet
from trac.env import Environment
from trac.resource import Resource
from trac.wiki.model import WikiPage
from trac.attachment import Attachment
from .helpers import set_req_keys, get_base_url
from .doc import Doc

#env = Environment('/home/user/Workspace/t11518/tracdev')
#resource = Resource('wiki', 'WikiStart', 1)
#page = WikiPage(env, resource)
#page.version

TEMPLATE_INSTANCE = 'req'
TEMPLATE_PAGE = 'Attachments'
TEMPLATE_NAME = 'template.docx'
#TEMPLATE_NAME = 'template.docm'

class WikiToDoc(Component):
    implements(INavigationContributor, ITemplateProvider, IRequestHandler)

    # INavigationContributor methods
    def get_active_navigation_item(self, req):
        return 'wiki2doc'

    def get_navigation_items(self, req):
        yield ('mainnav', 'wikitodoc',
               tag.a('Wiki to Doc', href=req.href.wiki2doc()))

    # IRequestHandler methods
    def match_request(self, req):
        """Each IRequestHandler is called to match a web request.
        The first matching handler is then used to process the request.
        Matching a request usually means checking the req.path_info
        string (the part of the URL relative to the Trac root URL)
        against a specific string prefix or regular expression.
        """
        return re.match(r'/wiki2doc(?:_trac)?(?:/.*)?$', req.path_info)

    def process_request(self, req):
        """Process the request. Return a (template_name, data) pair,
        where data is a dictionary of substitutions for the Jinja2
        template (the template context, in Jinja2 terms).

        Optionally, the return value can also be a (template_name, data,
        metadata) triple, where metadata is a dict with hints for the
        template engine or the web front-end."""

        print('DIR_req', dir(req))

        self.errorlog = []
        action = req.args.get('create_report', '__FORM_TOKEN')
        req_keys = set_req_keys(req)

        print('req keys', req_keys)
        print('action:', action)
        print('self.env', self.env)

        if req.method == 'POST':

            print('request is not:', req)
            print('request args:', req.args)

            page_path = req.args.get('get_wiki_link')

            print('page_path', page_path)

            match_path = re.match(
                r"(http://|e:)(.*|/)wiki/(.*)",
                page_path)

            if match_path:
                spec_name = re.split(r'\s+', match_path.group(3))
                spec_name = spec_name[0]
                spec_name = spec_name.split("|")
                spec_name = spec_name[0]
                spec_name = urllib.unquote(spec_name)
                print(spec_name)
                #resource = Resource('wiki', spec_name[0], 1)
                page = WikiPage(self.env, spec_name)

                print(page.name)

                self.process_report_task(page, req)
        else:
            pass

        data = {}
        add_stylesheet(req, 'hw/css/wiki2doc.css')
        # This tuple is for Genshi (template_name, data, content_type)
        # Without data the trac layout will not appear.
        return 'wiki2doc.html', data, None

    # ITemplateProvider methods
    # Used to add the plugin's templates and htdocs
    def get_templates_dirs(self):
        from pkg_resources import resource_filename
        return [resource_filename(__name__, 'templates')]

    def get_template(self, req):
        """ return path of standard auto report template """

        print("get_template:")
        print(req)
        
        page_path = get_base_url(req) + 'wiki/' + TEMPLATE_PAGE
#         self.envs[TEMPLATE_INSTANCE].project_url +\
#             '/wiki/' + TEMPLATE_PAGE

        print("page_path", page_path)

        for attachment in Attachment.select(self.env,
                                            'wiki',
                                            TEMPLATE_PAGE):
            if attachment.filename == TEMPLATE_NAME:
                return attachment.path
        self.errorlog.append(
            ("Attachment {} could not be found at {}.".\
             format(TEMPLATE_NAME, TEMPLATE_PAGE),
             0,
             page_path))

    def get_htdocs_dirs(self):
        """Return a list of directories with static resources (such as style
        sheets, images, etc.)

        Each item in the list must be a `(prefix, abspath)` tuple. The
        `prefix` part defines the path in the URL that requests to these
        resources are prefixed with.

        The `abspath` is the absolute path to the directory containing the
        resources on the local file system.
        """
        from pkg_resources import resource_filename
        return [('hw', resource_filename(__name__, 'htdocs'))]
    
    def process_report_task(self, page, req):
        """ process selected create apo and
            create report tasks.
            parameters = [sel_apo_tasks,
                          keys[4],
                          req]"""

        report = self.create_report(req)

        sections = self.get_sections_with_images(page, req)
        #sections = get_sections_with_tables(sections)
        
        print('sections', sections)

        #return self.errorlog, report.get_content()
    
    def create_report(self, req):
        """ Create SAR report """

        args = []

        print('self.get_template:', self.get_template(req))

        args = [self.get_template(req),
                self.env,
                self,
                req]
         
        report = Report(args)

        return report

    def get_sections_with_images(self, page, req):
        """ given a list of sections, returns a list of sections
            with attached images stored in a dictionary where key
            is the image file name in the spec and value is the
            file path to that image """

        sections_with_imgs = []
        spec_images = {}
        img_list = []
        path_list = []
        img_filename = None
        img_path = None
        image = re.compile(r'\s*\[\[Image\((.*(\.jpg|\.png|\.gif))\)\]\]\s*')

        text = page.text
        if text is not None:
            for line in text.splitlines():
                match = image.match(line)
                if match:
                    img_filename = match.group(1)
                    img_path = \
                        self.get_image_file(img_filename,
                                            page,
                                            req)
                    if img_filename and img_path:
                        img_list.append(img_filename)
                        path_list.append(img_path)
        spec_images = dict(zip(img_list, path_list))
        sections_with_imgs.append([text, spec_images])
        spec_images = {}

        return sections_with_imgs

    def get_image_file(self, filename, page, req):
        """ return path of image attachment """

        page_path = req.args.get('get_wiki_link')

        if page.exists:
            for attachment in Attachment.select(page.env,
                                                page.realm,
                                                page.resource.id):
                if attachment.filename == filename:
#                    path = str(attachment.path)
                    return attachment.path
            self.errorlog.append(
                ("Attachment {} could not be found at {}".\
                 format(filename, page.resource.id),
                 0,
                 page_path))
        else:
            self.errorlog.append(
                ("Page for the spec " +\
                 "{} could not be found!".format(page.name),
                 0,
                 page_path))

    