import os
import sys

sys.path.append(os.path.join(os.path.dirname(__file__), 'libs'))

import Foundation
import tempfile
from shutil import copyfile

import json

import adsk.cam
import adsk.core
import adsk.fusion
import traceback

USER_PARAM = 'exporter_stl'
COMMAND_KEY = 'ExportDesign'

keepHandlers = []
deletableObjects = []


def find_brep(chain, component, ui):
    if len(chain) > 1:
        if component.objectType == 'adsk::fusion::Component':
            child = component.occurrences.itemByName(chain[0])
        else:
            child = component.childOccurrences.itemByName(chain[0])
        return find_brep(chain[1:], child, ui)
    elif len(chain) == 1:
        return component.bRepBodies.itemByName(chain[0])
    return None


class ExportCommandCreatedEventHandler(adsk.core.CommandCreatedEventHandler):
    def notify(self, cmd):
        app = adsk.core.Application.get()
        ui = app.userInterface
        try:
            design = app.activeProduct
            dir_param = design.userParameters.itemByName(USER_PARAM)
            data = None
            if dir_param:
                data = json.loads(dir_param.comment)
            else:
                dialog = ui.createFileDialog()
                dialog.title = 'Select export file'
                dialog.filter = '*.*'
                accessible = dialog.showSave()
                if accessible == adsk.core.DialogResults.DialogOK:
                    file_name = dialog.filename
                    # we can take a URL only if the file exists, so we are doing a "touch"
                    with open(file_name, 'a'):
                        os.utime(file_name)
                    url = Foundation.NSURL.alloc().initFileURLWithPath_(file_name)
                    bookmark, error = url.bookmarkDataWithOptions_includingResourceValuesForKeys_relativeToURL_error_(
                        Foundation.NSURLBookmarkCreationWithSecurityScope,
                        None,
                        None,
                        None)
                    sys.stdout.flush()
                    # base64 encode the bookmark so it can be jsonified
                    the_bytes = bookmark.base64EncodedStringWithOptions_(0)
                    ui.messageBox('select body')
                    selected = ui.selectEntity('select body', 'Bodies,MeshBodies')
                    entity = selected.entity
                    chain = [entity.name]
                    obj = entity
                    while True:
                        obj = obj.assemblyContext
                        if obj:
                            chain.append(obj.name)
                        else:
                            break

                    data = {'file': the_bytes, 'chain': list(reversed(chain))}
                    json_params = json.dumps(data)
                    design.userParameters.add(USER_PARAM, adsk.core.ValueInput.createByString('0'), '', json_params)
            if data:
                body = find_brep(data['chain'], design.rootComponent, ui)
                if body:
                    nsdata = Foundation.NSData.alloc().initWithBase64EncodedString_options_(data['file'], 0)
                    url, is_stale, error = Foundation.NSURL.URLByResolvingBookmarkData_options_relativeToURL_bookmarkDataIsStale_error_(
                        nsdata,
                        Foundation.NSURLBookmarkResolutionWithSecurityScope,
                        None,
                        None,
                        None)
                    file_path = url.path()
                    accessible = url.startAccessingSecurityScopedResource()
                    if accessible:
                        try:
                            # using the temp dir, for some reasons exportManager tries to access the directory
                            # surrounding the output file
                            with tempfile.NamedTemporaryFile(suffix='.stl') as temp_file:
                                export_manager = design.exportManager
                                options = export_manager.createSTLExportOptions(body, temp_file.name)
                                options.sendToPrintUtility = False
                                options.meshRefinement = 0
                                export_manager.execute(options)
                                temp_file.seek(0)
                                copyfile(temp_file.name, file_path)
                            ui.messageBox('wrote STL file')
                        finally:
                            url.stopAccessingSecurityScopedResource()
                    else:
                        ui.messageBox('file ' + str(url) + ' not accessible!')
                else:
                    ui.messageBox('.'.join(data['chain']) + 'not found')
        except:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def run(context):
    app = adsk.core.Application.get()
    ui = app.userInterface
    try:
        print('STARTING')
        cmd_defs = ui.commandDefinitions
        my_button = ui.commandDefinitions.itemById(COMMAND_KEY)
        if my_button:
            my_button.deleteMe()
        my_button = cmd_defs.addButtonDefinition(COMMAND_KEY, 'Export STL', 'Export the pre-determined STL file',
                                                 'exportIcon')
        on_command_created = ExportCommandCreatedEventHandler()
        my_button.commandCreated.add(on_command_created)
        keepHandlers.append(on_command_created)
        deleter = replace_existing_control(ui.allToolbarPanels.itemById('SolidMakePanel').controls, my_button)
        deleter2 = replace_existing_control(ui.toolbars.itemById('QAT').controls, my_button)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def replace_existing_control(controls, new_command=None):
    control = controls.itemById(COMMAND_KEY)
    if control:
        control.deleteMe()
    if new_command:
        control = controls.addCommand(new_command)
        return lambda: control.deleteMe()
    return lambda: None


def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        button = ui.commandDefinitions.itemById(COMMAND_KEY)
        if button:
            button.deleteMe()
        replace_existing_control(ui.allToolbarPanels.itemById('SolidMakePanel').controls)
        global keepHandlers
        keepHandlers = []
        # for obj in deletableObjects:
        #    obj.deleteMe()
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
