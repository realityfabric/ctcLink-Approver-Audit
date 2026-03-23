# Approver Validation

## Assumptions

Purchasing Approvers match Expense Approvers 1:1. i.e.:
- Department Manager for Department XXXXX would also be the expense approver for Department XXXXX.
- The VP approver for XXXXX will also be the VP Expense Approver for XXXXX.
- Department Managers and VPs may need to approve Requisitions, Purchase Orders, and Vouchers.
- ZZ Expense Approval is dynamically applied to Expense Approvers.

## Preparation

You will need to run the following queries and save them as XLS[X]:
- ALL_DEPTS_BY_SETID
- QFS_SEC_EOAW_APPROVAL_SETUP
- QFS_SEC_OPR_EXP_APPRVR
- QFS_SEC_USER_ROLES_BY_UNIT

You can use QFS_DS_QUERY_RECORD_USER_RPT to determine which Query Security Roles are required to run the above queries.

## Road Map

- [ ] Security Checks
  - [X] Check for inactive employees with approval roles.
  - [ ] Check for Department Managers without the appropriate approval roles.
  - [ ] Check for employees with AWE routing roles without the appropriate approval roles.
    - [ ] ZZ_AW_EXEC_LEVEL_X (Pres & VPs)
    - [ ] Commodity Codes
    - [ ] Grants
    - [ ] Billing & AP Review
    - [ ] Amount Level roles
	- [ ] Route Control Configuration
- [ ] Travel & Expenses Approver Checks
  - [ ] Expense Approvers who are not Department Managers
  - [ ] Departments with no Expense Approver
  - [ ] Departments with an Expense Approver that is not the Department Manager
- [ ] Check for Inactive Employees who are Department Managers
- [ ] Check for Departments with no Department Manager
