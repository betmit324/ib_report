from my_logging import print_to_log
import datetime

debug_it = 0


def to_date(my_value):
    if isinstance(my_value, datetime.datetime):
        return my_value
    elif isinstance(my_value, int):
        print_to_log("Converting " + str(my_value) + " to date")
        return datetime.datetime.strptime("Jan 1, 1900", "%b %d, %Y") + datetime.timedelta(days=my_value)
    else:
        try:
            return datetime.datetime.strptime(my_value, "%b %d, %Y")
        except (ValueError, TypeError):
            print_to_log("Could not convert '" + str(my_value) + "' to date")
            return ''


class HardwareContract:
    # Betsy working on skip_entitlements_check for these products 5/3
    products_to_exclude = []
    response_profiles_list = []

    def __init__(self, service_level,
                 warranty_end_date,
                 service_type,
                 response_profile,
                 service_contract_id,
                 contract_end_date,
                 months_till_expire,
                 #removed from SAM -contract_status,
                 hw_service_level_status,
                 sc_entitlement_status,
                 serial=''):

        self.service_level = service_level
        self.warranty_end_date = warranty_end_date
        self.service_type = service_type
        self.response_profile = response_profile
        self.service_contract_id = service_contract_id
        self.contract_end_date = contract_end_date
        self.serial = serial
        try:
            self.months_till_expire = int(months_till_expire)
        except (ValueError, TypeError):
            if debug_it:
                print_to_log("Unexpected months_till_expire value: " + str(months_till_expire))
                print_to_log("Using 0 for months_till_expire for serial " + str(self.serial))
            self.months_till_expire = 0
            #removed from SAM -self.contract_status = contract_status
        self.hw_service_level_status = hw_service_level_status
        self.sc_entitlement_status = sc_entitlement_status

    def __bool__(self):
        try:
            if self.hw_service_level_status and len(self.hw_service_level_status) > 0:
                return True
        except TypeError:
            print_to_log("Strange hw service level status for serial " + str(self.serial))
        try:
            if self.response_profile and len(self.response_profile) > 0:
                return True
        except TypeError:
            print_to_log("Strange response profile for serial " + str(self.serial))
        return False

    def __eq__(self, other):

        try:
            if self.hw_service_level_status == other.hw_service_level_status and self.response_profile == other.response_profile and self.months_till_expire == other.months_till_expire:
                return True
            else:
                return False
        except TypeError:
            return False

    def __ne__(self, other):

        try:
            if self.hw_service_level_status == other.hw_service_level_status and self.response_profile == other.response_profile and self.months_till_expire == other.months_till_expire:
                return False
            else:
                return True
        except TypeError:
            return True

    def __str__(self):

        if self.hw_service_level_status:
            string = self.hw_service_level_status
        else:
            string = "Blank"

        if self.response_profile:
            string += " " + self.response_profile
        else:
            string += " Blank"

        if self.months_till_expire:
            string += " " + str(self.months_till_expire)
        else:
            string += " Blank"

        return string

    def __and__(self, other):
        # Compares two HW contracts and returns the greater entitlement
        if not (other.months_till_expire or other.hw_service_level_status or other.response_profile):
            return self
        if not (self.months_till_expire or self.hw_service_level_status or self.response_profile):
            return other

        if self.months_till_expire is not None and self.months_till_expire > 0 and (other.months_till_expire is None or other.months_till_expire <= 0):
            return self
        if other.months_till_expire is not None and other.months_till_expire > 0 and (self.months_till_expire is None or self.months_till_expire <= 0):
            return other
        # if self.hw_service_level_status == "Expired" and other.hw_service_level_status == "":
        if self.hw_service_level_status and other.hw_service_level_status == "":
            return self
        # if other.hw_service_level_status == "Expired" and self.hw_service_level_status == "":
        if other.hw_service_level_status and self.hw_service_level_status == "":
            return other
        if (self.months_till_expire > 0 and other.months_till_expire > 0) or (self.months_till_expire <= 0 and other.months_till_expire <= 0):
            try:
                if HardwareContract.response_profiles_list.index(self.response_profile) < HardwareContract.response_profiles_list.index(other.response_profile):
                    return self
                elif HardwareContract.response_profiles_list.index(self.response_profile) > HardwareContract.response_profiles_list.index(other.response_profile):
                    return other
                elif HardwareContract.response_profiles_list.index(self.response_profile) == HardwareContract.response_profiles_list.index(other.response_profile):
                    try:
                        if self.months_till_expire > other.months_till_expire:
                            return self
                        else:
                            return other
                    except TypeError:
                        print_to_log("Problem comparing contracts, unexpected 'months till expire', contact neil.maldonado@netapp.com with serial " + str(self.serial))
                        return self

            except ValueError:
                string = "Problem comparing contracts 1) "
                if self.service_type:
                    string += self.service_type
                if self.service_contract_id:
                    string += self.service_contract_id + " "
                if self.response_profile:
                    string += self.response_profile + " "
                string += " 2) "
                if other.service_contract_id:
                    string += other.service_contract_id + " "
                if other.service_type:
                    string += other.service_type
                if other.response_profile:
                    string += other.response_profile
                print_to_log(string + ", contact neil.maldonado@netapp.com with serial " + str(self.serial))
                return self
        else:
            print_to_log("Contract compare failed, contact neil.maldonado@netapp.com with serial " + str(self.serial))
            print_to_log(str(self))
            print_to_log(str(other))
            return self

    def __or__(self, other):
        # Compares two HW contracts and returns the longest entitlement
        if not (other.months_till_expire or other.hw_service_level_status or other.response_profile):
            return self
        if not (self.months_till_expire or self.hw_service_level_status or self.response_profile):
            return other

        if self.months_till_expire is not None and self.months_till_expire > 0 and (other.months_till_expire is None or other.months_till_expire <= 0):
            return self
        if other.months_till_expire is not None and other.months_till_expire > 0 and (self.months_till_expire is None or self.months_till_expire <= 0):
            return other
        # if self.hw_service_level_status == "Expired" and other.hw_service_level_status == "":
        if self.hw_service_level_status and other.hw_service_level_status == "":
            return self
        # if other.hw_service_level_status == "Expired" and self.hw_service_level_status == "":
        if other.hw_service_level_status and self.hw_service_level_status == "":
            return other
        if self.months_till_expire > other.months_till_expire:
            # print_to_log("contract 'or' returning " + str(self.months_till_expire) + " months instead of " + str(other.months_till_expire))
            return self
        elif self.months_till_expire < other.months_till_expire:
            # print_to_log("contract 'or' returning " + str(other.months_till_expire) + " months instead of " + str(self.months_till_expire))
            return other
        elif self.months_till_expire == other.months_till_expire:
            try:
                if HardwareContract.response_profiles_list.index(self.response_profile) < HardwareContract.response_profiles_list.index(other.response_profile):
                    return self
                elif HardwareContract.response_profiles_list.index(self.response_profile) > HardwareContract.response_profiles_list.index(other.response_profile):
                    return other
                elif HardwareContract.response_profiles_list.index(self.response_profile) == HardwareContract.response_profiles_list.index(other.response_profile):
                    print_to_log("Duplicate contracts found for same serial, contact neil.maldonado@netapp.com with serial " + str(self.serial))
                    return self

            except ValueError:
                string = "Problem comparing contracts 1) "
                if self.service_type:
                    string += self.service_type
                if self.service_contract_id:
                    string += self.service_contract_id + " "
                if self.response_profile:
                    string += self.response_profile + " "
                string += " 2) "
                if other.service_contract_id:
                    string += other.service_contract_id + " "
                if other.service_type:
                    string += other.service_type
                if other.response_profile:
                    string += other.response_profile
                print_to_log(string + ", contact neil.maldonado@netapp.com with serial " + str(self.serial))
                return self
        else:
            print_to_log("Contract compare failed, contact neil.maldonado@netapp.com with serial " + str(self.serial))
            print_to_log(str(self))
            print_to_log(str(other))
            return self


class NetAppSystem:
    failed_attributes = []

    def __init__(self, serial, ib_row=0):
        self.ib_row = ib_row
        self.owner = ""
        self.site = ""
        self.group = ""
        self.solution = ""
        self.serial = serial
        self.partner_serial = ""
        self.hostname = ""
        self.cluster_name = ""
        self.cluster_serial = ""
        self.cluster_uuid = ""
        self.os_version = ""
        self.asup_status = ""
        self.declined = ""
        self.asup_date = ""
        self.product_family = ""
        self.platform = ""
        self.controller_eos_date = ""
        self.pvr_flag = ""
        self.pvr_date = ""
        self.first_eos_date = ""
        self.first_eos_part = ""
        self.age = ""
        self.ha_pair_flag = ""
        self.service_level = ""
        self.sc_service_level = ""
        self.entitlement_status = ""
        self.sc_entitlement_status = ""
        self.contact_name = ""
        self.contact_number = ""
        self.contact_email = ""
        self.raw_tb = ""
        self.num_shelves = ""
        self.num_disks = ""
        self.service_level = ""
        self.warranty_end_date = ""
        self.service_type = ""
        self.response_profile = ""
        self.service_contract_id = ""
        self.contract_end_date = ""
        self.months_till_expire = ""
        #removed from SAM -self.contract_status = ""
        self.hw_service_level_status = ""
        self.nrd_service_contract_id = ""
        self.nrd_contract_end_date = ""
        self.nrd_months_till_expire = ""
        self.nrd_hw_service_level_status = ""
        self.notes = []
        # IB Team fields
        self.account_name = ""
        self.end_customer = ""
        self.customer_geo = ""
        self.customer_country = ""
        self.customer_city = ""
        self.original_so = ""
        self.shipped_date = ""
        self.software_service_end_date = ""
        self.reseller = ""
        self.solidfire_tag = ""
        self.old_serial = ""
        self.filer_use = ""
        self.flash_list = []
        self.longest_service_level = ""
        self.longest_warranty_end_date = ""
        self.longest_service_type = ""
        self.longest_response_profile = ""
        self.longest_service_contract_id = ""
        self.longest_contract_end_date = ""
        self.longest_months_till_expire = ""
        #removed from SAM -self.longest_contract_status = ""
        self.longest_hw_service_level_status = ""
        self.longest_sc_entitlement_status = ""

    def __str__(self):
        temp_csv = ','.join([
            self.owner,
            self.site,
            self.group,
            self.solution,
            self.serial,
            self.hostname,
            self.cluster_name,
            self.cluster_serial,
            self.cluster_uuid,
            self.os_version,
            self.asup_status,
            self.declined,
            str(self.asup_date),
            self.product_family,
            self.platform,
            str(self.controller_eos_date),
            self.pvr_flag,
            str(self.pvr_date),
            str(self.first_eos_date),
            self.first_eos_part,
            str(self.age),
            self.ha_pair_flag,
            self.service_level,
            self.sc_service_level,
            self.entitlement_status,
            self.sc_entitlement_status,
            self.contact_name,
            self.contact_number,
            self.contact_email,
            str(self.raw_tb),
            str(self.num_shelves),
            str(self.num_disks),
            self.service_level,
            str(self.warranty_end_date),
            self.service_type,
            self.response_profile,
            str(self.service_contract_id),
            str(self.contract_end_date),
            str(self.months_till_expire),
            #removed from SAM -self.contract_status,
            self.hw_service_level_status,
            self.account_name,
            self.end_customer,
            self.customer_geo,
            self.customer_country,
            self.customer_city,
            self.original_so,
            self.shipped_date,
            self.software_service_end_date,
            self.reseller,
            self.solidfire_tag,
            self.old_serial,
            self.filer_use]
        )

        if self.notes:
            temp_csv += "," + self.notes[-1].replace(',', ';')
        return temp_csv

    def list_changes(self, other):
        changes_list = []
        attributes_to_ignore = ['ib_row', 'raw_tb', 'asup_date', 'age', 'months_till_expire', 'notes', 'nrd_months_till_expire', 'flash_list', 'longest_months_till_expire']
        for k, v in vars(self).items():
            try:
                if k not in attributes_to_ignore and other.__getattribute__(k) != v:
                    try:
                        if v.date() != other.__getattribute__(k).date():
                            changes_list.append((k, v, other.__getattribute__(k)))
                    except (AttributeError, ValueError):
                        if other.__getattribute__(k) != v:
                            changes_list.append((k, v, other.__getattribute__(k)))
            except AttributeError:
                if k not in NetAppSystem.failed_attributes:
                    NetAppSystem.failed_attributes.append(k)
                    print_to_log("Problem comparing " + k + " to previous week, possibly a new column added to the report")

        return changes_list

    def set_owner(self, my_value):
        if my_value:
            self.owner = my_value
        else:
            self.owner = ""

    def set_site(self, my_value):
        if my_value:
            self.site = my_value
        else:
            self.site = ""

    def set_group(self, my_value):
        if my_value:
            self.group = my_value
        else:
            self.group = ""

    def set_solution(self, my_value):
        if my_value:
            self.solution = my_value
        else:
            self.solution = ""

    def set_hostname(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.hostname = ""
        elif my_value:
            self.hostname = my_value
        else:
            self.hostname = ""

    def set_cluster_name(self, my_value):
        if my_value:
            self.cluster_name = my_value
        else:
            self.cluster_name = ""

    def set_cluster_serial(self, my_value):
        if my_value:
            self.cluster_serial = my_value
        else:
            self.cluster_serial = ""

    def set_cluster_uuid(self, my_value):
        if my_value:
            self.cluster_uuid = my_value
        else:
            self.cluster_uuid = ""

    def set_os_version(self, my_value):
        if my_value:
            self.os_version = my_value
        else:
            self.os_version = ""

    def set_asup_status(self, my_value):
        if my_value:
            self.asup_status = my_value
        else:
            self.asup_status = ""

    def set_declined(self, my_value):
        if my_value:
            self.declined = my_value
        else:
            self.declined = ""

    def set_asup_date(self, my_value):
        if my_value:
            self.asup_date = to_date(my_value)
        else:
            self.asup_date = ""

    def set_product_family(self, my_value):
        if my_value:
            self.product_family = my_value
        else:
            self.product_family = ""

    def set_platform(self, my_value):
        if my_value:
            self.platform = my_value
        else:
            self.platform = ""

    def set_controller_eos_date(self, my_value):
        if my_value:
            self.controller_eos_date = to_date(my_value)
        else:
            self.controller_eos_date = ""

    def set_pvr_flag(self, my_value):
        if my_value:
            self.pvr_flag = my_value
        else:
            self.pvr_flag = ""

    def set_pvr_date(self, my_value):
        if my_value:
            self.pvr_date = to_date(my_value)
        else:
            self.pvr_date = ""

    def set_first_eos_date(self, my_value):
        if my_value:
            self.first_eos_date = to_date(my_value)
        else:
            self.first_eos_date = ""

    def set_first_eos_part(self, my_value):
        if my_value:
            self.first_eos_part = my_value
        else:
            self.first_eos_part = ""

    def set_age(self, my_value):
        if my_value:
            self.age = my_value
        else:
            self.age = ""

    def set_ha_pair_flag(self, my_value):
        if my_value:
            self.ha_pair_flag = my_value
        else:
            self.ha_pair_flag = ""

    # Betsy working on skip_entitlements_check for these products 5/3
    def set_service_level(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.service_level = "entitlement not needed"
        elif my_value:
            self.service_level = my_value
        else:
            self.service_level = ""
    # Betsy working on skip_entitlements_check for these products 5/3
    def set_sc_service_level(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.sc_service_level = "entitlement not needed"
        elif my_value:
            self.sc_service_level = my_value
        else:
            self.sc_service_level = ""

    # Betsy working on skip_entitlements_check for these products 5/3
    def set_entitlement_status(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.entitlement_status = "entitlement not needed"
        elif my_value:
            self.entitlement_status = my_value
        else:
            self.entitlement_status = ""

    # Betsy working on skip_entitlements_check for these products 5/3
    def set_sc_entitlement_status(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.sc_entitlement_status = "entitlement not needed"
        elif my_value:
            self.sc_entitlement_status = my_value
        else:
            self.sc_entitlement_status = ""

    def set_contact_name(self, my_value):
        if my_value:
            self.contact_name = my_value
        else:
            self.contact_name = ""

    def set_contact_number(self, my_value):
        if my_value:
            self.contact_number = my_value
        else:
            self.contact_number = ""

    def set_contact_email(self, my_value):
        if my_value:
            self.contact_email = my_value
        else:
            self.contact_email = ""

    def set_raw_tb(self, my_value):
        if my_value:
            self.raw_tb = my_value
        else:
            self.raw_tb = ""

    def set_num_shelves(self, my_value):
        if my_value:
            self.num_shelves = my_value
        else:
            self.num_shelves = ""

    def set_num_disks(self, my_value):
        if my_value:
            self.num_disks = my_value
        else:
            self.num_disks = ""

    def set_warranty_end_date(self, my_value):
        if my_value:
            self.warranty_end_date = to_date(my_value)
        else:
            self.warranty_end_date = ""

    def set_service_type(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.service_type = "no entitlement needed"
        elif my_value:
            self.service_type = my_value
        else:
            self.service_type = ""

    def set_response_profile(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.response_profile = "no entitlement needed"
        elif my_value:
            self.response_profile = my_value
        else:
            self.response_profile = ""

    def set_service_contract_id(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.service_contract_id = "no entitlement needed"
        elif my_value:
            self.service_contract_id = my_value
        else:
            self.service_contract_id = ""

    def set_contract_end_date(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.contract_end_date = "no entitlement needed"
        elif my_value:
            self.contract_end_date = to_date(my_value)
        else:
            self.contract_end_date = ""

    def set_months_till_expire(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.months_till_expire = ""
        elif my_value:
            if type(my_value) != int:
                try:
                    self.months_till_expire = int(my_value)
                except ValueError:
                    print_to_log("Unexpected months_till_expire value: " + str(my_value))
                    print_to_log("Using 0 for months_till_expire")
                    self.months_till_expire = 0
            else:
                self.months_till_expire = my_value
        else:
            self.months_till_expire = 0

            #removed from SAM - def set_contract_status(self, my_value):
            #removed from SAM -if self.product_family in HardwareContract.products_to_exclude:
            #removed from SAM - self.contract_status = "no entitlement needed"
            #removed from SAM - elif my_value:
            #removed from SAM -  self.contract_status = my_value
            #removed from SAM - else:
            #removed from SAM - self.contract_status = ""

    def set_hw_service_level_status(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.hw_service_level_status = "no entitlement needed"
        elif my_value:
            self.hw_service_level_status = my_value
        else:
            self.hw_service_level_status = ""

    def set_nrd_service_contract_id(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.nrd_service_contract_id = "no entitlement needed"
        elif my_value:
            self.nrd_service_contract_id = my_value
        else:
            self.nrd_service_contract_id = ""

    def set_nrd_contract_end_date(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.nrd_contract_end_date = ""
        elif my_value:
            self.nrd_contract_end_date = to_date(my_value)
        else:
            self.nrd_contract_end_date = ""

    def set_nrd_months_till_expire(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.nrd_months_till_expire = ""
        elif my_value:
            if type(my_value) != int:
                try:
                    self.nrd_months_till_expire = int(my_value)
                except ValueError:
                    print_to_log("Unexpected nrd_months_till_expire value: " + str(my_value))
                    print_to_log("Using 0 for nrd_months_till_expire")
                    self.nrd_months_till_expire = 0
            else:
                self.nrd_months_till_expire = my_value
        else:
            self.nrd_months_till_expire = 0

    def set_nrd_hw_service_level_status(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.nrd_hw_service_level_status = ""
        elif my_value:
            self.nrd_hw_service_level_status = my_value
        else:
            self.nrd_hw_service_level_status = ""

    def add_note(self, my_value):
        self.notes.append(my_value)

    def set_partner_serial(self, my_value):
        if my_value:
            self.partner_serial = my_value
        else:
            self.partner_serial = ""

    def set_account_name(self, my_value):
        if my_value:
            self.account_name = my_value
        else:
            self.account_name = ""

    def set_end_customer(self, my_value):
        if my_value:
            self.end_customer = my_value
        else:
            self.end_customer = ""

    def set_customer_geo(self, my_value):
        if my_value:
            self.customer_geo = my_value
        else:
            self.customer_geo = ""

    def set_customer_country(self, my_value):
        if my_value:
            self.customer_country = my_value
        else:
            self.customer_country = ""

    def set_customer_city(self, my_value):
        if my_value:
            self.customer_city = my_value
        else:
            self.customer_city = ""

    def set_original_so(self, my_value):
        if my_value:
            self.original_so = my_value
        else:
            self.original_so = ""

    def set_shipped_date(self, my_value):
        if my_value:
            self.shipped_date = to_date(my_value)
        else:
            self.shipped_date = ""

    def set_software_service_end_date(self, my_value):
        if my_value:
            self.software_service_end_date = to_date(my_value)
        else:
            self.software_service_end_date = ""

    def set_reseller(self, my_value):
        if my_value:
            self.reseller = my_value
        else:
            self.reseller = ""

    def set_solidfire_tag(self, my_value):
        if my_value:
            self.solidfire_tag = my_value
        else:
            self.solidfire_tag = ""

    def set_old_serial(self, my_value):
        if my_value:
            self.old_serial = my_value
        else:
            self.old_serial = ""

    def set_filer_use(self, my_value):
        if my_value:
            self.filer_use = my_value
        else:
            self.filer_use = ""

    def set_longest_service_level(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.longest_service_level = ""
        elif my_value:
            self.longest_service_level = my_value
        else:
            self.longest_service_level = ""

    def set_longest_warranty_end_date(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.longest_warranty_end_date = ""
        elif my_value:
            self.longest_warranty_end_date = to_date(my_value)
        else:
            self.longest_warranty_end_date = ""

    def set_longest_service_type(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.longest_service_type = ""
        elif my_value:
            self.longest_service_type = my_value
        else:
            self.longest_service_type = ""

    def set_longest_response_profile(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.longest_response_profile = ""
        elif my_value:
            self.longest_response_profile = my_value
        else:
            self.longest_response_profile = ""

    def set_longest_service_contract_id(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.longest_service_contract_id = ""
        elif my_value:
            self.longest_service_contract_id = my_value
        else:
            self.longest_service_contract_id = ""

    def set_longest_contract_end_date(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.longest_contract_end_date = ""
        elif my_value:
            self.longest_contract_end_date = to_date(my_value)
        else:
            self.longest_contract_end_date = ""

    def set_longest_months_till_expire(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.longest_months_till_expire = ""
        elif my_value:
            if type(my_value) != int:
                try:
                    self.longest_months_till_expire = int(my_value)
                except ValueError:
                    print_to_log("Unexpected months_till_expire value: " + str(my_value))
                    print_to_log("Using 0 for months_till_expire")
                    self.longest_months_till_expire = 0
            else:
                self.longest_months_till_expire = my_value
        else:
            self.longest_months_till_expire = 0

        #removed from SAM - def set_longest_contract_status(self, my_value):
            #removed from SAM -if self.product_family in HardwareContract.products_to_exclude:
            #removed from SAM - self.longest_contract_status = ""
            #removed from SAM - elif my_value:
            #removed from SAM -   self.longest_contract_status = my_value
            #removed from SAM - else:
            #removed from SAM -  self.longest_contract_status = ""

    def set_longest_hw_service_level_status(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.longest_hw_service_level_status = ""
        elif my_value:
            self.longest_hw_service_level_status = my_value
        else:
            self.longest_hw_service_level_status = ""

    def set_longest_sc_entitlement_status(self, my_value):
        if self.product_family in HardwareContract.products_to_exclude:
            self.longest_sc_entitlement_status = ""
        elif my_value:
            self.longest_sc_entitlement_status = my_value
        else:
            self.longest_sc_entitlement_status = ""
