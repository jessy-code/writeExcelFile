#!/usr/bin/python

from __future__ import (absolute_import, division, print_function)

__metaclass__ = type

DOCUMENTATION = r'''
---
module: writeExcelFileLine

short_description: Write a line in an Excel file.

version_added: "1.0.0"

description: Write a line in an Excel file. Create the file if it doesn't exist.

options:
    dest:
        description: Complete file path where the Excel file will be saved
        required: true
        type: str
    sheet:
        description: Sheet name target
        required: true
        type: str
    line_number:
        description: The line number target
        required: true
        type: int
    line:
        description: The line content
        required: true
        type: list of string

author:
    - Jessy Martin (@jessy-code)
'''

EXAMPLES = r'''
# write the line list in the 3thd line of file /tmp/test.xlsx
- name: Write a line in an Excel file
  writeExcelFileLine:
    dest: '/tmp/test.xlsx'
    sheet: 'sheet'
    line_number: 3
    line:
      - test1
      - test2
      - test1
      - test4
'''

RETURN = r'''
# These are examples of possible return values, and in general should use other names for return values.
original_message:
    description: The original name param that was passed in.
    type: str
    returned: always
    sample: 'hello world'
message:
    description: The output message that the test module generates.
    type: str
    returned: always
    sample: 'goodbye'
'''

from ansible.module_utils.basic import AnsibleModule
from ansible.module_utils.ExcelFile import ExcelFile


def run_module():
    # define available arguments/parameters a user can pass to the module
    module_args = dict(
        dest=dict(type='str', required=True),
        sheet=dict(type='str', required=True),
        line_number=dict(type='int', required=False, default=1),
        start_column=dict(type='int', required=False, default=1),
        line=dict(type='list', required=True)
    )

    # seed the result dict in the object
    # we primarily care about changed and state
    # changed is if this module effectively modified the target
    # state will include any data that you want your module to pass back
    # for consumption, for example, in a subsequent task
    result = dict(
        changed=False,
        original_message='',
        message=''
    )

    # the AnsibleModule object will be our abstraction working with Ansible
    # this includes instantiation, a couple of common attr would be the
    # args/params passed to the execution, as well as if the module
    # supports check mode
    module = AnsibleModule(
        argument_spec=module_args,
        supports_check_mode=True
    )

    # if the user is working with this module in only check mode we do not
    # want to make any changes to the environment, just return the current
    # state with no modifications
    if module.check_mode:
        module.exit_json(**result)

    # manipulate or modify the state as needed (this is going to be the
    # part where your module will do what it needs to do)
    # result['original_message'] = module.params['name']
    # result['message'] = 'goodbye'

    # use whatever logic you need to determine whether or not this module
    # made any modifications to your target
    # if module.params['new']:
    #    result['changed'] = True

    # during the execution of the module, if there is an exception or a
    # conditional state that effectively causes a failure, run
    # AnsibleModule.fail_json() to pass in the message and the result
    # if module.params['name'] == 'fail me':
    #    module.fail_json(msg='You requested this to fail', **result)

    dest = module.params['dest']
    exce_file = ExcelFile(dest)
    exce_file.write_list_in_line_of_sheet(module.params['sheet'], module.params['line_number'],
                                          module.params['line'], module.params['start_column'])

    # in the event of a successful module execution, you will want to
    # simple AnsibleModule.exit_json(), passing the key/value results
    module.exit_json(**result)


def main():
    run_module()


if __name__ == '__main__':
    main()
