# Approver Validation

## Assumptions

Purchasing Approvers match Expense Approvers 1:1. i.e.:
- Department Manager for Department XXXXX would also be the expense approver for Department XXXXX.
- The VP approver for XXXXX will also be the VP Expense Approver for XXXXX.

## Preparation

You will need to run the following queries and save them as XLS[X]:
- ALL_DEPTS_BY_SETID
- QFS_SEC_EOAW_APPROVAL_SETUP
- QFS_SEC_OPR_EXP_APPRVR
- QFS_SEC_USER_ROLES_BY_UNIT

You can use QFS_DS_QUERY_RECORD_USER_RPT to determine which Query Security Roles are required to run the above queries.