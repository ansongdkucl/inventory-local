#!/usr/bin/env python3

import argparse
import os
from pathlib import Path

import yaml
from dotenv import load_dotenv
from nornir import InitNornir
from nornir.core.inventory import ConnectionOptions
from nornir.core.task import Task, Result


# ============================================================
# SETTINGS
# ============================================================

CONNECTION_OPTIONS_FILE = Path("connection_options.yaml")

DEFAULT_NETMIKO_EXTRAS = {
    "timeout": 30,
    "session_timeout": 30,
    "conn_timeout": 30,
    "banner_timeout": 30,
    "auth_timeout": 30,
    "fast_cli": False,
    "global_delay_factor": 2,
}


PLATFORM_MAP = {
    # Cisco
    "cisco": "cisco_ios",
    "ios": "cisco_ios",
    "cisco_ios": "cisco_ios",
    "cisco_os": "cisco_ios",
    "cisco_os_ios": "cisco_ios",

    # Aruba OS-Switch / ProCurve
    "aruba": "aruba_osswitch",
    "aruba_os": "aruba_osswitch",
    "aruba_osswitch": "aruba_osswitch",
    "hp_procurve": "hp_procurve",
    "procurve": "hp_procurve",

    # Aruba CX
    "aoscx": "aruba_aoscx",
    "aruba_cx": "aruba_aoscx",
    "aruba_aoscx": "aruba_aoscx",
}


# ============================================================
# HELPERS
# ============================================================

def normalise_platform(platform: str) -> str:
    """
    Converts inventory platform names into Netmiko device_type names.
    Example:
        cisco_os      -> cisco_ios
        aruba_os      -> aruba_osswitch
        aruba_aoscx   -> aruba_aoscx
    """
    if not platform:
        return "cisco_ios"

    cleaned = platform.strip().lower().replace("-", "_")
    return PLATFORM_MAP.get(cleaned, cleaned)


def load_netmiko_extras() -> dict:
    """
    Loads optional Netmiko connection options from connection_options.yaml.
    """
    extras = DEFAULT_NETMIKO_EXTRAS.copy()

    if not CONNECTION_OPTIONS_FILE.exists():
        return extras

    with open(CONNECTION_OPTIONS_FILE, "r", encoding="utf-8") as file:
        data = yaml.safe_load(file) or {}

    file_extras = data.get("netmiko", {}) or {}
    extras.update(file_extras)

    return extras


def get_username() -> str:
    return (
        os.environ.get("NORNIR_USERNAME")
        or os.environ.get("USERNAME")
        or os.environ.get("username")
    )


def get_password(platform: str) -> str:
    """
    Chooses the correct password based on platform.

    Aruba CX and Aruba OS-Switch use the Aruba password.
    Cisco uses the Cisco password.
    """
    if "aruba" in platform or "aoscx" in platform:
        return (
            os.environ.get("NORNIR_ARUBA_PASSWORD")
            or os.environ.get("passwordAD")
            or os.environ.get("PASSWORD")
            or os.environ.get("password")
        )

    return (
        os.environ.get("NORNIR_CISCO_PASSWORD")
        or os.environ.get("PASSWORD")
        or os.environ.get("password")
    )


def get_secret(platform: str) -> str | None:
    """
    Cisco enable secret only.
    """
    if "cisco" not in platform:
        return None

    return (
        os.environ.get("NORNIR_CISCO_SECRET")
        or os.environ.get("secret")
    )


def host_in_group(host, group_name: str) -> bool:
    """
    Safely checks if a host is in a Nornir group.
    """
    for group in host.groups:
        if getattr(group, "name", str(group)) == group_name:
            return True

    return False


# ============================================================
# DEVICE CONNECTION
# ============================================================

def setup_device_connection(task: Task) -> Result:
    host = task.host

    platform = normalise_platform(host.platform)
    username = get_username()
    password = get_password(platform)
    secret = get_secret(platform)

    if not username or not password:
        return Result(
            host=host,
            failed=True,
            result="Missing username/password. Check your .env file.",
        )

    extras = load_netmiko_extras()

    if secret:
        extras["secret"] = secret

    # Update the host platform so Netmiko receives the corrected device_type
    host.platform = platform

    host.connection_options["netmiko"] = ConnectionOptions(
        hostname=host.hostname,
        port=host.port or 22,
        username=username,
        password=password,
        platform=platform,
        extras=extras,
    )

    return Result(
        host=host,
        result=f"Connection options prepared. platform={platform}",
    )


# ============================================================
# CONNECTION TEST TASK
# ============================================================

def connection_test(task: Task) -> Result:
    setup_result = setup_device_connection(task)

    if setup_result.failed:
        return setup_result

    try:
        connection = task.host.get_connection(
            "netmiko",
            task.nornir.config,
        )

        prompt = connection.find_prompt()

        return Result(
            host=task.host,
            result=f"SSH OK | platform={task.host.platform} | prompt={prompt}",
        )

    except Exception as error:
        return Result(
            host=task.host,
            failed=True,
            result=f"SSH FAIL | platform={task.host.platform} | error={error}",
        )


# ============================================================
# MAIN
# ============================================================

def main():
    load_dotenv()

    parser = argparse.ArgumentParser(
        description="Nornir connection test for testbed devices"
    )

    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Nornir config file. Default: config.yaml",
    )

    parser.add_argument(
        "--group",
        default="testbed",
        help="Nornir group to test. Default: testbed",
    )

    parser.add_argument(
        "--host",
        help="Test one host by inventory name or IP address",
    )

    args = parser.parse_args()

    nr = InitNornir(config_file=args.config)

    if args.host:
        nr = nr.filter(
            filter_func=lambda host: (
                host.name == args.host
                or host.hostname == args.host
            )
        )
    else:
        nr = nr.filter(
            filter_func=lambda host: host_in_group(host, args.group)
        )

    if len(nr.inventory.hosts) == 0:
        print("No devices selected.")
        print(f"Check that your host.yaml has devices in group: {args.group}")
        return

    print(f"Testing {len(nr.inventory.hosts)} device(s)...\n")

    results = nr.run(
        task=connection_test,
        raise_on_error=False,
    )

    passed = 0
    failed = 0

    for hostname, multi_result in results.items():
        result = multi_result[0]

        if result.failed:
            failed += 1
            print(f"[FAIL] {hostname} - {result.result}")
        else:
            passed += 1
            print(f"[OK]   {hostname} - {result.result}")

    nr.close_connections(on_good=True, on_failed=True)

    print("\nSummary")
    print("-------")
    print(f"Passed: {passed}")
    print(f"Failed: {failed}")


if __name__ == "__main__":
    main()