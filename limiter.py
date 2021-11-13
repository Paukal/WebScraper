import win32job
import win32process
import win32security

def limit_memory(megabyte_limit=100):
    "Create some obscure Wandows object necessary for setting memory limits."

    # Create job object
    sec = win32security.SECURITY_ATTRIBUTES()
    job = win32job.CreateJobObject(sec, "LimitedMemoryPythonJob")

    # Set memory limits on job object
    limits = win32job.QueryInformationJobObject(job,
        win32job.JobObjectExtendedLimitInformation)
    limits['JobMemoryLimit'] = megabyte_limit * 1024 ** 2
    limits['BasicLimitInformation']['LimitFlags'] |= \
        win32job.JOB_OBJECT_LIMIT_JOB_MEMORY
    win32job.SetInformationJobObject(job,
        win32job.JobObjectExtendedLimitInformation, limits)

    # Assign current process to job object
    process = win32process.GetCurrentProcess()
    win32job.AssignProcessToJobObject(job, process)
